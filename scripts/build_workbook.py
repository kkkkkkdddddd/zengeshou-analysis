
"""Recreate the local insurance analysis workbook from a verified transcription."""
import argparse
import csv
import itertools
import math
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

HEADERS = (
    "保单年度", "年末年龄", "当期保费", "身故保险金", "现金价值", "年金",
    "年度红利", "红利+年金", "预期生存总利益", "总利益增长额", "总利益年增长率",
)
RAW = "'原图数据'!"
YEAR = "'逐年核算'!"
CASH = "'现金流与实际利率'!"
EXIT = "'逐年退保利率'!"
SUMMARY = "'分段分析'!"
RC = 7
TC = 15
INVALID_IRR = "需复核现金流"


@dataclass(frozen=True)
class Record:
    cells: tuple

    @property
    def age(self):
        return self.cells[1]

    @property
    def premium(self):
        return self.event(2)

    @property
    def annuity(self):
        return self.event(5)

    @property
    def dividend(self):
        return self.event(6)

    def event(self, index):
        value = self.cells[index]
        return value if isinstance(value, (int, float)) else 0


def normalize_header(value):
    text = re.sub(r"\s+", "", str(value)).replace("＋", "+")
    return "年末年龄" if text == "年龄年末" else text


def rows_from(path, sheet_name):
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as src:
            yield from csv.reader(src)
        return
    if path.suffix.lower() != ".xlsx":
        raise ValueError("图片须先按模板转录。脚本输入为 CSV 或 XLSX。")
    from openpyxl import load_workbook

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        tab = book[sheet_name] if sheet_name else (
            book["原图数据"] if "原图数据" in book.sheetnames else book.active
        )
        yield from tab.values
    finally:
        book.close()


def parse_number(value, *, percent=False, allow_text=False):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, bool):
        raise ValueError("布尔值不能作为保单数值")
    if isinstance(value, str):
        clean = re.sub(r"[\s,，]", "", value).replace("％", "%").replace("−", "-")
        is_pct = clean.endswith("%")
        try:
            amount = float(clean[:-1] if is_pct else clean)
        except ValueError:
            if allow_text:
                return value
            raise ValueError(f"无法解析数值 {value!r}") from None
        if is_pct and not percent:
            raise ValueError("金额列出现了百分比")
        value = amount / 100 if is_pct else amount
    if not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError("源表含无效数值")
    return int(value) if int(value) == value else value


def read_records(path, sheet_name, blank_events_zero):
    source = iter(rows_from(path, sheet_name))
    if not any(tuple(map(normalize_header, row[:11])) == HEADERS for row in source):
        raise ValueError("找不到 11 列表头，请参照空白模板")
    records = []
    finished = False
    for row in source:
        listed = (list(row[:11]) + [None] * 11)[:11]
        if all(v is None or v == "" for v in listed):
            continue
        try:
            year = parse_number(listed[0])
        except ValueError:
            if listed[1] is None or listed[1] == "":
                finished = True
                continue
            raise
        if finished:
            raise ValueError("表格在注释之后又出现了数据行，请把脚注移到表末")
        try:
            cells = tuple(parse_number(v, percent=i == 10, allow_text=i in (2, 9, 10))
                          for i, v in enumerate(listed))
            age = cells[1]
            if year != len(records) + 1 or not isinstance(age, int):
                raise ValueError("保单年度须从 1 连续编号，年末年龄须为整数")
            if records and age != records[-1].age + 1:
                raise ValueError("年末年龄有重复或断行")
            for idx in (2, 5, 6):
                amount = cells[idx]
                if not isinstance(amount, (int, float)) and not blank_events_zero:
                    raise ValueError(f"{HEADERS[idx]}为空或为文字，确认无收付后可使用 --blank-events-zero")
            if any(isinstance(cells[i], (int, float)) and cells[i] < 0 for i in range(2, 9)):
                raise ValueError("原表的保费、给付及价值列应录入非负金额")
        except ValueError as exc:
            raise ValueError(f"源表第 {len(records)+1} 个保单年度：{exc}") from None
        records.append(Record(cells))
    if not records or not any(rec.premium for rec in records):
        raise ValueError("至少需要一条年度记录及一笔保费")
    return records


def ledger(records, index, with_dividend, surrender):
    flows = [0] * (index + 2)
    for t, rec in enumerate(records[:index+1], 1):
        flows[t-1] -= rec.premium
        flows[t] += rec.annuity + (rec.dividend if with_dividend else 0)
    if surrender:
        flows[-1] += records[index].cells[4]
    return flows


def present_value(flows, rate):
    discount = 1 / (1 + rate)
    accumulated = 0.0
    for amount in reversed(flows):
        accumulated = amount + discount * accumulated
    return accumulated


def irr(flows):
    signs = [amount > 0 for amount in flows if amount]
    if not signs or signs[0] or not signs[-1] or signs != sorted(signs):
        return None
    low, high = -0.999999, 1.0
    while present_value(flows, high) > 0 and high < 1e12:
        high = high * 2 + 1
    if not (present_value(flows, low) > 0 >= present_value(flows, high)):
        return None
    for _ in range(90):
        middle = (low + high) / 2
        if present_value(flows, middle) > 0:
            low = middle
        else:
            high = middle
    return (low + high) / 2


class Output:
    def __init__(self, book, cents):
        self.amount = book.add_format({"num_format": "#,##0.00" if cents else "#,##0"})
        self.rate = book.add_format({"num_format": "0.00%"})
        self.cents = book.add_format({"num_format": "#,##0.00"})
        self.wrap = book.add_format({"text_wrap": True, "valign": "vcenter"})

    def fx(self, sheet, r, c, formula, value, fmt=None):
        sheet.write_formula(r-1, c-1, "=" + formula, fmt, "" if value is None else value)

    def rate_cell(self, sheet, r, c, range_ref, flows):
        rate = irr(flows)
        if rate is None:
            sheet.write(r-1, c-1, INVALID_IRR)
        else:
            self.fx(sheet, r, c, f"IRR({range_ref},{rate:.15g})", rate, self.rate)
        return rate


def configure(sheet, widths, frozen=0):
    for i, width in enumerate(widths):
        sheet.set_column(i, i, width)
    if frozen:
        sheet.freeze_panes(frozen, 0)


def write_raw(out, sheet, records, args):
    configure(sheet, [12,12,14,16,14,13,14,15,19,17,19], 6)
    sheet.activate()
    sheet.write("A1", args.title)
    sheet.write("A2", "用户提供的保险利益演示表。保留原始金额、百分比、空白和文字。单位为人民币元。")
    sheet.write("A3", "公式入口：现金流与实际利率（B5 改退保年龄）、逐年退保利率（各年份）。")
    sheet.write_row("A6", HEADERS)
    for r, rec in enumerate(records, RC):
        for c, value in enumerate(rec.cells):
            if value is None:
                continue
            fmt = out.rate if c == 10 else out.amount if 2 <= c < 10 else None
            if isinstance(value, str):
                stop = r
                if c in (9, 10) and ("缴费" in value or "交费" in value):
                    while stop < RC + len(records) - 1 and records[stop-RC+1].cells[c] is None:
                        stop += 1
                if stop > r:
                    sheet.merge_range(r-1, c, stop-1, c, "", out.wrap)
                    sheet.write_string(r-1, c, value, out.wrap)
                else:
                    sheet.write_string(r-1, c, value, out.wrap if "\n" in value else None)
                    if "\n" in value:
                        sheet.set_row(r-1, 31)
            else:
                sheet.write_number(r-1, c, value, fmt)
    for i, note in enumerate(args.note, RC+len(records)+1):
        sheet.write_string(i-1, 0, note)


def write_years(out, sheet, records, args, hold_index):
    configure(sheet, [12,14,14,15,16,18,17,18,20,16,18,17,16,16,16], 6)
    sheet.write("A1", "逐年核算")
    sheet.write("A2", "演示差额＝演示总利益－年末现金价值－累计年金－累计红利。金额以分项为准。")
    sheet.write("A3", "达标列按全表保费的两倍计算。退出合计或累计领取按每年领出相应演示红利。")
    sheet.write("A4", "滚存假设")
    if args.roll_rate is not None:
        sheet.write_number("B4", args.roll_rate, out.rate)
    sheet.write_row("A6", [
        "年末年龄","累计保费","累计年金","累计演示红利","年末现金价值","原图预期总利益",
        "演示差额","复算年增长额","增长额÷累计保费","总利益同比","滚存复算总利益","原图减复算",
        "演示达两倍","领出加退保达两倍","仅领出达两倍"
    ])
    premium = ann = div = 0
    target = 2*sum(rec.premium for rec in records)
    first = [None]*3
    for i, rec in enumerate(records):
        r = RC + i
        premium, ann, div = premium+rec.premium, ann+rec.annuity, div+rec.dividend
        cv, combined, total = (rec.cells[j] for j in (4,7,8))
        out.fx(sheet,r,1,f"{RAW}B{r}",rec.age)
        for col, rc, value in ((2,"C",premium),(3,"F",ann),(4,"G",div)):
            out.fx(sheet,r,col,f"SUM({RAW}{rc}{RC}:{rc}{r})",value,out.amount)
        for col, rc, value in ((5,"E",cv),(6,"I",total)):
            out.fx(sheet,r,col,f'IF(ISNUMBER({RAW}{rc}{r}),{RAW}{rc}{r},"")',value,out.amount)
        if cv is not None and total is not None:
            out.fx(sheet,r,7,f"F{r}-SUM(C{r}:E{r})",total-cv-ann-div,out.amount)
        if i:
            prev = r-1
            previous_cv, previous_total = (records[i-1].cells[j] for j in (4,8))
            if previous_total is not None and total is not None:
                delta = total - previous_total
                out.fx(sheet,r,8,f"F{r}-F{prev}",delta,out.amount)
                if premium:
                    out.fx(sheet,r,9,f"H{r}/B{r}",delta/premium,out.rate)
                if previous_total:
                    out.fx(sheet,r,10,f"H{r}/F{prev}",delta/previous_total,out.rate)
                if args.roll_rate is not None and None not in (previous_cv,cv,combined):
                    rolled = (previous_total-previous_cv)*(1+args.roll_rate)+combined+cv
                    out.fx(sheet,r,11,f"(F{prev}-E{prev})*(1+$B$4)+{RAW}H{r}+E{r}",rolled,out.cents)
                    out.fx(sheet,r,12,f"F{r}-K{r}",total-rolled,out.cents)
        for k, (amount, expr) in enumerate((
            (total,f"F{r}"),
            (ann+div+cv if cv is not None and i <= hold_index
             and not (args.payout_age and i == hold_index) else None,f"SUM(C{r}:E{r})"),
            (ann+div if i <= hold_index else None,f"SUM(C{r}:D{r})"),
        )):
            if amount is not None:
                hit = int(amount>=target)
                out.fx(sheet,r,13+k,f"--({expr}>={SUMMARY}B4)",hit)
                if hit and first[k] is None:
                    first[k] = rec.age
    return first


def write_summary(out, sheet, records, args, terminal, first_ages):
    configure(sheet,[25,14,14,17,17,17,18,20], 0)
    sheet.write("A1", "分段分析")
    sheet.write("A2", "阶段按保费、年金和末次给付划分，红利采用图示金额。")
    target = 2*sum(rec.premium for rec in records)
    last = RC+len(records)-1
    sheet.write("A4","名义翻倍门槛（元）")
    out.fx(sheet,4,2,f"2*SUM({RAW}C{RC}:C{last})",target,out.amount)
    sheet.write_row("A6", ["翻倍口径","首次年龄"])
    options = ["海报演示总利益","领出加退保现金价值","仅累计领出"]
    for r, (label, col, age) in enumerate(zip(options, ("M","N","O"), first_ages), 7):
        sheet.write(r-1,0,label)
        flags = f"{YEAR}{col}{RC}:{col}{last}"
        out.fx(sheet,r,2,f'IF(COUNTIF({flags},1),INDEX({YEAR}A{RC}:A{last},MATCH(1,{flags},0)),"未达到（表内）")',
               age if age is not None else "未达到（表内）")
    sheet.write("A11","后两项采用图示红利逐年领出。退出合计使用对应年末现金价值。")
    sheet.write_row("A13", [
        "阶段","起始年龄","结束年龄","阶段交费","阶段年金","阶段演示红利","期末现金价值","期末预期总利益",
    ])
    last_premium = max(i for i,x in enumerate(records) if x.premium)
    first_annuity = next((i for i,x in enumerate(records) if x.annuity),None)

    def label(i):
        if args.payout_age and i >= terminal:
            return "末次给付及之后"
        if i <= last_premium:
            return "缴费阶段"
        if first_annuity is None or i < first_annuity:
            return "等待或观察"
        return "首年领取" if i == first_annuity else "逐年领取"

    for r,(stage,group) in enumerate(itertools.groupby(range(len(records)), label),14):
        indexes = list(group)
        start,end = indexes[0],indexes[-1]
        a,b = RC+start,RC+end
        sheet.write_row(r-1,0,[stage,records[start].age,records[end].age])
        for c,(source_col,attr) in enumerate((("C","premium"),("F","annuity"),("G","dividend")),4):
            value = sum(getattr(records[i],attr) for i in indexes)
            out.fx(sheet,r,c,f"SUM({RAW}{source_col}{a}:{source_col}{b})",value,out.amount)
        for c,(source_col,idx) in enumerate((("E",4),("I",8)),7):
            value = records[end].cells[idx]
            out.fx(sheet,r,c,f'IF(ISNUMBER({RAW}{source_col}{b}),{RAW}{source_col}{b},"")',value,out.amount)


def write_cash(out, sheet, records, args, terminal, candidates, selected):
    configure(sheet,[32,14,16,18,16,16,17,21,21,22,23],14)
    sheet.write("A1","按逐年收付计算保险年化")
    sheet.write("A2","时点 0 交首笔保费。年度末领款与下一年度初交费落在同一时点。金额单位为人民币元。")
    sheet.write("A3","年末退出在领取当年款项后计入一次现金价值，停止后续交费。")
    sheet.write("A4","含红利情景按图示红利逐年领出。红利按零情景保留其他图示金额。")
    sheet.write("A5","退保年龄（键入数字）")
    last_candidate_row = 6+len(candidates)
    if selected is not None:
        sheet.write_number("B5",records[selected].age)
        sheet.data_validation("B5",{
            "validate":"integer","criteria":"between",
            "minimum":records[candidates[0]].age,"maximum":records[candidates[-1]].age,
            "error_title":"需要数值年龄","error_message":"请输入可测算的整数年龄",
        })
        out.fx(sheet,5,7,f'IF(ISNUMBER(B5),IF(COUNTIF({EXIT}A7:A{last_candidate_row},B5)=1,"","请输入有效的数值年龄"),"请输入有效的数值年龄")',"")
    else:
        sheet.write("G5","观察期无可计算退保年龄")
    sheet.write_row("A6",["测算情形","交费合计","年金及红利","退保金","名义净额","年化IRR","现值校验（元）"])
    sheet.write("A12","现值校验把每笔现金流按 IRR 折回投保日。数值应接近零。")
    sheet.write_row("A14",[
        "距投保年数","年龄节点","时点计划保费","图示年金","演示年度红利","图示现金价值",
        "选定年龄退保金","退保净流（红利0）","退保净流（含红利）","终点净流（红利0）","终点净流（含红利）",
    ])
    flows = [ledger(records,selected,div,True) if selected is not None else []
             for div in (False,True)] + [ledger(records,terminal,div,args.payout_age is None)
                                        for div in (False,True)]
    final_r = TC+terminal+1
    for t in range(len(records)+1):
        r=TC+t
        age = records[t-1].age if t else records[0].age-1
        sheet.write_row(r-1,0,[t,age])
        if t<len(records):
            out.fx(sheet,r,3,f"-SUM({RAW}C{RC+t})",-records[t].premium,out.amount)
        else:
            sheet.write_number(r-1,2,0,out.amount)
        ann, div, cv = (records[t-1].annuity,records[t-1].dividend,records[t-1].cells[4]) if t else (0,0,None)
        if t:
            for col, rc, val in ((4,"F",ann),(5,"G",div)):
                out.fx(sheet,r,col,f"SUM({RAW}{rc}{RC+t-1})",val,out.amount)
            out.fx(sheet,r,6,f'IF(ISNUMBER({RAW}E{RC+t-1}),{RAW}E{RC+t-1},"")',cv,out.amount)
        else:
            sheet.write_row(r-1,3,[0,0],out.amount)
        terminal_add = f"+F{r}" if args.payout_age is None else ""
        selected_pay = cv if selected is not None and t == selected+1 else 0
        out.fx(sheet,r,7,f'IF(AND($G$5="",B{r}=$B$5),F{r},0)',selected_pay,out.amount)
        for j,expr in enumerate((
            f'IF($G$5<>"",0,IF(B{r}<$B$5,C{r}+D{r},IF(B{r}=$B$5,D{r}+G{r},0)))',
            f'H{r}+IF(AND($G$5="",B{r}<=$B$5),E{r},0)',
            f'IF(A{r}<{terminal+1},C{r}+D{r},IF(A{r}={terminal+1},D{r}{terminal_add},0))',
            f'J{r}+IF(A{r}<={terminal+1},E{r},0)',
        )):
            out.fx(sheet,r,8+j,expr,flows[j][t] if t<len(flows[j]) else 0,out.amount)

    last_cash = TC+len(records)
    terminal_label = f"生存至{records[terminal].age}岁" if args.payout_age else f"{records[terminal].age}岁退出"
    for j,(name,col) in enumerate(zip(("选定年龄退保","选定年龄退保",terminal_label,terminal_label),"HIJK")):
        r = 7+j
        div = j % 2
        sheet.write(r-1,0,name + ("（含演示红利）" if div else "（红利按0）"))
        index = selected if j < 2 else terminal
        if index is None:
            continue
        values = flows[j]
        observations = records[:index+1]
        amount_premium = sum(x.premium for x in observations)
        amount_in = sum(x.annuity+(x.dividend if div else 0) for x in observations)
        cv = records[index].cells[4] if j<2 or args.payout_age is None else 0
        range_ref = f"{col}{TC}:{col}{last_cash if j<2 else final_r}"
        selected_guard = lambda expression: f'IF($G$5="",{expression},"")'
        if j < 2:
            pay = f'-SUMIF(B{TC}:B{last_cash},"<"&$B$5,C{TC}:C{last_cash})'
            receive = f'SUMIF(B{TC}:B{last_cash},"<="&$B$5,D{TC}:D{last_cash})'
            if div:
                receive += f'+SUMIF(B{TC}:B{last_cash},"<="&$B$5,E{TC}:E{last_cash})'
            for c,(expr,val) in enumerate(((pay,amount_premium),(receive,amount_in),
                                            (f"SUM(G{TC}:G{last_cash})",cv),(f"SUM({range_ref})",sum(values))),2):
                out.fx(sheet,r,c,selected_guard(expr),val,out.amount)
            column = "G" if div else "F"
            rate = irr(values)
            lookup = f"INDEX({EXIT}{column}7:{column}{last_candidate_row},MATCH($B$5,{EXIT}A7:A{last_candidate_row},0))"
            out.fx(sheet,r,6,selected_guard(lookup),rate if rate is not None else INVALID_IRR,out.rate)
        else:
            for c,(expr,val) in enumerate((
                (f"-SUM(C{TC}:C{final_r-1})",amount_premium),
                (f"SUM(D{TC}:D{final_r})"+(f"+SUM(E{TC}:E{final_r})" if div else ""),amount_in),
                (f"F{final_r}" if args.payout_age is None else "0",cv),
                (f"SUM({range_ref})",sum(values)),
            ),2):
                out.fx(sheet,r,c,expr,val,out.amount)
            rate = out.rate_cell(sheet,r,6,range_ref,values)
        resid = present_value(values,rate) if rate is not None else None
        tail = last_cash if j<2 else final_r
        out.fx(sheet,r,7,f'IF(ISNUMBER(F{r}),{col}{TC}+NPV(F{r},{col}{TC+1}:{col}{tail}),"")',resid,out.cents)


def write_exits(out, sheet, records, candidates, terminal):
    configure(sheet,[13,13,16,15,14,16,18,4],6)
    sheet.write("A1","逐年退保利率")
    sheet.write("A2","逐行按该年退出。只有原图列示数值现金价值的年龄参与计算。")
    sheet.write("A3","向右查看逐笔现金流矩阵，分别列红利按零和含演示红利。")
    sheet.write_row("A6",["退保年龄","保单年度","退出现金价值","已交保费","累计年金","红利按零IRR","含演示红利IRR"])
    n = terminal+2
    for j,label in enumerate(("红利按零现金流","含演示红利现金流")):
        start=8+j*n
        sheet.write(4,start,label)
        sheet.set_column(start,start+n-1,13)
        sheet.write_row(5,start,[records[0].age-1]+[x.age for x in records[:terminal+1]])
    for r,index in enumerate(candidates,7):
        rec=records[index]
        terminal_t=index+1
        cr=TC+terminal_t
        sheet.write_row(r-1,0,[rec.age,terminal_t])
        for c,(expr,value) in enumerate((
            (f"{CASH}F{cr}",rec.cells[4]),
            (f"-SUM({CASH}C{TC}:C{cr-1})",sum(x.premium for x in records[:index+1])),
            (f"SUM({CASH}D{TC}:D{cr})",sum(x.annuity for x in records[:index+1])),
        ),3):
            out.fx(sheet,r,c,expr,value,out.amount)
        for div in (False,True):
            start=8+div*n
            values=ledger(records,index,div,True)
            for t in range(n):
                source=TC+t
                h=xl_col_to_name(start+t)
                flow=f"{CASH}{'K' if div else 'J'}{source}"
                payout=f"{CASH}D{source}"+(f"+{CASH}E{source}" if div else "")
                out.fx(sheet,r,start+t+1,
                       f"IF({h}$6<$A{r},{flow},IF({h}$6=$A{r},{payout}+$C{r},0))",
                       values[t] if t<len(values) else 0,out.amount)
            ref=f"{xl_col_to_name(start)}{r}:{xl_col_to_name(start+n-1)}{r}"
            out.rate_cell(sheet,r,7 if div else 6,ref,values)


def write_glossary(sheet, out):
    configure(sheet,[22,50,47,54],3)
    sheet.write("A1","术语解释")
    sheet.write_row("A3",["术语","白话","本表怎么读","什么条件下涉及钱"])
    terms=[
        ("保单年度／年末年龄","以保单周年为界，年龄列指当年度末。","先对照投保年龄与原表年度","具体以保单生效日和领取日为准。"),
        ("当期保费","当年度交给保险公司的钱。","模型列在该年度初","提前退出时不再计后续保费。"),
        ("身故保险金","发生符合条件的身故时给受益人的钱。","原表独立列出","与生存给付、退保分别判断。"),
        ("现金价值","保单在指定时点的价值，常用来计算退保金。","只在退出当年入账一次","核对同年已领款和实际退保扣减。"),
        ("年金／年度红利","年金按约给付。红利是当年分配的浮动金额。","逐年分别录入并测算有无红利","需满足领取条件，未来红利可能为零。"),
        ("红利+年金","同一年度两项金额的合计。","原图合计保留，IRR 使用分项","实际按各自给付约定处理。"),
        ("预期生存总利益","演示中的生存权益合计，有时还包含此前给付滚存。","对照现金价值、累计给付和差额","需核对实际到账方式及滚存条件。"),
        ("总利益增长额／年增长率","相邻演示金额的差和比例。","图示比例可能使用已交保费作分母","统计指标，现金流列另列实际收付。"),
        ("年化IRR","把逐年的交费和领取折到同一时点后相抵的利率。","默认为按年间隔的复利年化","取决于收付时点、退出年龄和红利。"),
    ]
    for r,term in enumerate(terms,4):
        sheet.write_row(r-1,0,term,out.wrap)
        sheet.set_row(r-1,32)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input",type=Path,help="经核对的 CSV 或 XLSX 原始数据")
    parser.add_argument("--output",type=Path,required=True,help="生成的 XLSX")
    parser.add_argument("--sheet",help="指定 XLSX 原始工作表")
    parser.add_argument("--title",default="保险利益演示原表")
    parser.add_argument("--note",action="append",default=[])
    parser.add_argument("--payout-age",type=int,help="明确的末次生存给付年龄")
    parser.add_argument("--exit-age",type=int,help="默认退保年龄")
    parser.add_argument("--roll-rate",type=float,help="待检验的滚存比例，例如 0.023")
    parser.add_argument("--blank-events-zero",action="store_true")
    args=parser.parse_args()
    try:
        if args.output.suffix.lower()!=".xlsx" or args.output.resolve()==args.input.resolve():
            raise ValueError("输出须为另一个 .xlsx 文件")
        if args.roll_rate is not None and (not math.isfinite(args.roll_rate) or args.roll_rate<=-1):
            raise ValueError("滚存假设须为大于 -1 的有限数值比例")
        records=read_records(args.input,args.sheet,args.blank_events_zero)
        ages={rec.age:i for i,rec in enumerate(records)}
        if args.payout_age is not None:
            if args.payout_age not in ages:
                raise ValueError("末次给付年龄不在原表中")
            terminal=ages[args.payout_age]
        else:
            terminal=next((i for i in reversed(range(len(records))) if records[i].cells[4] is not None),None)
            if terminal is None:
                raise ValueError("原表未列现金价值。应从正式资料确认，或指定末次生存给付年龄")
        if 8+2*(terminal+2)>16384:
            raise ValueError("年度现金流矩阵超过 Excel 可用列数")
        limit=terminal if args.payout_age is not None else terminal+1
        candidates=[i for i in range(limit) if records[i].cells[4] is not None]
        if args.exit_age is not None:
            selected=ages.get(args.exit_age)
            if selected not in candidates:
                raise ValueError("退保年龄须具有数值现金价值并位于测算区间")
        else:
            selected=next((i for i in candidates if records[i].annuity),candidates[-1] if candidates else None)

        args.output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=args.output.parent,suffix=".xlsx",delete=False) as handle:
            staged=Path(handle.name)
        try:
            options={"strings_to_formulas":False,"strings_to_urls":False}
            with xlsxwriter.Workbook(staged,options) as book:
                raw,summary,year,glossary,cash,exits=(book.add_worksheet(name) for name in (
                    "原图数据","分段分析","逐年核算","术语解释","现金流与实际利率","逐年退保利率",
                ))
                has_cents=any(isinstance(v,float) for rec in records for v in rec.cells[2:10])
                out=Output(book,has_cents)
                write_raw(out,raw,records,args)
                first=write_years(out,year,records,args,terminal)
                write_summary(out,summary,records,args,terminal,first)
                write_glossary(glossary,out)
                write_cash(out,cash,records,args,terminal,candidates,selected)
                write_exits(out,exits,records,candidates,terminal)
            staged.replace(args.output)
        finally:
            staged.unlink(missing_ok=True)
        print(f"已生成 {len(records)} 个保单年度的工作簿。")
    except (ValueError,OSError,KeyError) as exc:
        parser.error(str(exc))


if __name__=="__main__":
    main()
