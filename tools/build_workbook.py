# -*- coding: utf-8 -*-
"""三経路のAPI認証とCognito M2M設定の解説ブックを作る。"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = ROOT / "API認証_三経路とCognito_M2M.xlsx"

INK = "1A1A1A"
CHARCOAL = "2C2C2C"
SLATE = "3F3F3F"
MID = "5C5C5C"
LINE = "C6C6C6"
PAPER = "F7F7F7"
WHITE = "FFFFFF"
BAND = "ECECEC"
ZEBRA = "F3F3F3"
INPUT_BG = "E4E4E4"

THIN = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
THICK = Border(
    left=Side(style="medium", color=MID),
    right=Side(style="medium", color=MID),
    top=Side(style="medium", color=MID),
    bottom=Side(style="medium", color=MID),
)


def font(size=10, bold=False, color=INK):
    return Font(name="Meiryo UI", size=size, bold=bold, color=color, scheme=None)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def align(horizontal="left"):
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=True)


def weight(text):
    total = 0
    for ch in str(text):
        total += 2 if ord(ch) > 127 else 1
    return total


def lines_for(text, excel_width):
    capacity = max(6, int(float(excel_width) * 0.78))
    total = 0
    for para in str(text).split("\n"):
        if para == "":
            total += 1
            continue
        units = weight(para)
        total += max(1, (units + capacity - 1) // capacity)
    return max(1, total)


def height_for(text, excel_width, size=10, minimum=22, cap=160):
    lines = lines_for(text, excel_width)
    return min(cap, max(minimum, 8 + lines * (size + 6)))


def apply_page(ws, title):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.45, right=0.45, top=0.7, bottom=0.55, header=0.26, footer=0.26
    )
    ws.sheet_properties.tabColor = MID
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_view.view = "normal"
    ws.oddHeader.left.text = '&"Meiryo UI"&9API認証 三経路と Cognito M2M'
    ws.oddHeader.right.text = '&"Meiryo UI"&9確認日 2026-10-14'
    ws.oddFooter.left.text = (
        '&"Meiryo UI"&8灰色の太枠は入力。白と薄灰の数値は数式。単価は米国東部の例。'
    )
    ws.oddFooter.right.text = '&"Meiryo UI"&9&P / &N'
    ws.oddFooter.center.text = f'&"Meiryo UI"&9{title}'
    ws.sheet_format.defaultRowHeight = 18
    ws.print_options.horizontalCentered = True
    ws.page_setup.horizontalDpi = 300
    ws.page_setup.verticalDpi = 300
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.page_setup.scale = 70
    ws.sheet_view.zoomScale = 110


def paint(cell, value, *, size=10, bold=False, color=INK, bg=None, horizontal="left", border=True, input_cell=False):
    cell.value = value
    cell.font = font(size, bold, color)
    cell.alignment = align(horizontal)
    if bg:
        cell.fill = fill(bg)
    cell.border = THICK if input_cell else (THIN if border else Border())
    cell.protection = Protection(locked=not input_cell)
    return cell


def set_widths(ws, widths):
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for col in range(len(widths) + 1, 16):
        ws.column_dimensions[get_column_letter(col)].width = 3
    ws.freeze_panes = "A3"
    ws.print_title_rows = "1:2"
    ws.page_setup.fitToWidth = 1
    ws.auto_filter.ref = None
    ws.sheet_view.zoomScale = 110


def merge_title(ws, cols, text, fill_color=CHARCOAL, size=16, font_color=WHITE):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    cell = ws.cell(1, 1, text)
    cell.font = font(size, True, font_color)
    cell.alignment = align("left")
    cell.fill = fill(fill_color)
    for col in range(1, cols + 1):
        other = ws.cell(1, col)
        other.fill = fill(fill_color)
        other.border = THIN
        other.font = font(size, True, font_color)
    ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    sub = ws.cell(2, 1, "確認日 2026-10-14。詳しい文章は同じフォルダの Markdown。灰色の太枠だけを書き換える。")
    sub.font = font(9, False, INK)
    sub.alignment = align("left")
    sub.fill = fill(BAND)
    for col in range(1, cols + 1):
        other = ws.cell(2, col)
        other.fill = fill(BAND)
        other.border = THIN
        other.font = font(9, False, INK)
    ws.row_dimensions[2].height = 22


def write_header(ws, row, headers, widths):
    for col, text in enumerate(headers, start=1):
        paint(ws.cell(row, col), text, size=10, bold=True, color=WHITE, bg=CHARCOAL, horizontal="center")
    ws.row_dimensions[row].height = 28
    ws.auto_filter.ref = None
    return row + 1


def write_rows(ws, row, rows, widths, zebra=True):
    for r_index, values in enumerate(rows):
        bg = ZEBRA if zebra and r_index % 2 else WHITE
        max_h = 22
        for col, value in enumerate(values, start=1):
            paint(ws.cell(row, col), value, bg=bg, horizontal="left" if col > 1 else "left")
            width = widths[col - 1]
            max_h = max(max_h, height_for(value, width, minimum=24, cap=150))
        ws.row_dimensions[row].height = max_h
        row += 1
    return row


def style_merged_rest(ws, row, cols, bg, size=10, bold=False, color=INK):
    for col in range(2, cols + 1):
        other = ws.cell(row, col)
        other.font = font(size, bold, color)
        other.fill = fill(bg)
        other.border = THIN
        other.alignment = align("left")


def section(ws, row, cols, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    paint(ws.cell(row, 1), text, size=12, bold=True, color=WHITE, bg=SLATE)
    style_merged_rest(ws, row, cols, SLATE, size=12, bold=True, color=WHITE)
    ws.row_dimensions[row].height = 24
    return row + 1


def note(ws, row, cols, text, widths):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    paint(ws.cell(row, 1), text, size=9, bg=PAPER)
    style_merged_rest(ws, row, cols, PAPER, size=9)
    total = sum(widths)
    ws.row_dimensions[row].height = height_for(text, total, size=9, minimum=32, cap=90)
    return row + 1


def build():
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)

    build_conclusion(wb.active)
    build_paths(wb.create_sheet("01_三つの入口"))
    build_issuer(wb.create_sheet("02_誰が発行するか"))
    build_reuse(wb.create_sheet("03_再利用と有効期間"))
    build_settings(wb.create_sheet("04_Cognitoの各設定"))
    build_motion(wb.create_sheet("05_動作イメージ"))
    build_history(wb.create_sheet("06_歴史"))
    build_perf(wb.create_sheet("07_性能の計算"))
    build_inputs(wb.create_sheet("08_調整つまみ"))
    build_sources(wb.create_sheet("09_出典"))

    named = {
        "AccessMin": "'08_調整つまみ'!$C$5",
        "IdMin": "'08_調整つまみ'!$C$6",
        "RefreshDays": "'08_調整つまみ'!$C$7",
        "SessionMin": "'08_調整つまみ'!$C$8",
        "SkewSec": "'08_調整つまみ'!$C$9",
        "SfClients": "'08_調整つまみ'!$C$10",
        "NetClients": "'08_調整つまみ'!$C$11",
        "BizRps": "'08_調整つまみ'!$C$12",
        "Hours": "'08_調整つまみ'!$C$13",
        "UnitPrice": "'08_調整つまみ'!$C$14",
        "CacheSec": "'08_調整つまみ'!$C$15",
    }
    for name, ref in named.items():
        wb.defined_names.add(DefinedName(name=name, attr_text=ref))

    default_font = Font(name="Meiryo UI", size=10, color=INK, scheme=None)
    wb._fonts[0].name = "Meiryo UI"
    wb._fonts[0].sz = 10
    wb._fonts[0].scheme = None
    wb.properties.title = "API認証 三経路と Cognito M2M"
    wb.properties.creator = "API認証の解説"
    wb.properties.subject = "イントラネットのAPI-Key、SalesforceとインターネットのOAuth、Cognito M2M設定"
    wb.properties.description = "確認日 2026-10-14。Meiryo UI、モノトーン。灰色の太枠が入力。"
    wb.properties.keywords = "Cognito,M2M,OAuth,API-Key"
    wb.active = 0

    # 余白のスタイル漏れを防ぐ。空セルは触らない。
    _ = default_font
    wb.save(XLSX_PATH)
    return XLSX_PATH


def build_conclusion(ws):
    ws.title = "00_結論"
    widths = [24, 28, 42, 55, 36]
    apply_page(ws, "結論")
    set_widths(ws, widths)
    merge_title(ws, 5, "API認証の結論。M2M（機械同士）を前提にする")
    row = section(ws, 3, 5, "五つの問いに対する答え")
    row = write_header(ws, row, ["問い", "結論", "誰が何をするか", "動作の要点", "性能試験"], widths)
    rows = [
        [
            "イントラネットAPI",
            "ヘッダの API-Key で認証する。アクセストークンは無い。",
            "運用側が合言葉を事前に渡す。呼び出し元が毎回ヘッダへ入れる。受け側が比較する。Cognito は発行しない。",
            "社内の呼び出し元 → イントラAPI用ALB → ECS。公開のトークン売場へは行かない。",
            "固定ヘッダを付ける。トークン取得の時間は測定に入らない。",
        ],
        [
            "Salesforce API",
            "OAuth 2.0 のクライアントクレデンシャルズである。仕様の「API認可トークンの取得API」はそのトークン要求の名前である。",
            "依頼は Salesforce。発行は Cognito ユーザープール。検証はプライベート API Gateway の Cognito オーソライザー。業務の ECS は発行しない。",
            "先に POST /oauth2/token。その後、Bearer を付けて Private Connect 経由で業務API。旅行は二本。",
            "対象に残るなら、クライアントごとにトークンを使い回す。毎リクエストの取得にはしない。",
        ],
        [
            "インターネットAPI",
            "業務呼び出しは同じ OAuth である。一覧にトークン取得の行が無くても、窓口は Cognito 上に実在する。",
            "依頼は外部システム。発行は同じ Cognito。検証はリージョナル API Gateway。窓口のホストは業務APIのホストとは別である。",
            "教室の一覧に、校門の外の切符売場を書かないだけである。売場のパスは /oauth2/token。2026-08-31 以降は GetClientToken という別窓口もある。",
            "対象外にしてよい。公開経路と、その呼び出し元のトークン要求は負荷から外れる。隠れた大量発行にはならない。",
        ],
        [
            "再利用と有効期間",
            "アクセストークンは期限まで再利用できる。提示設定では発行から 60分。expires_in は 3600秒。",
            "呼び出し元がメモリに保管し、Bearer として何度でも提示する。期限の前に同じ ID とシークレットで新規発行する。",
            "更新トークンは M2M の応答に入らない。30日での延長は無い。前倒しの初期値は 300秒。",
            "取り直しは約 55分に 1回。業務トランザクションの外で測る。",
        ],
        [
            "Cognito の各設定",
            "M2M の通行証に効くのは、クライアントID、シークレット、アクセストークン 60分である。",
            "30日、IDトークン 1時間、セッション 3分、取消、ユーザ存在エラーの防止は、人のログインの欄である。同じ画面に並ぶ。",
            "取消を有効にしても、発行済みの M2M アクセストークンは署名検証だけでは exp まで残る。",
            "インターネットを外しても、この読みは変わらない。イントラネットの測定はこれらのタイマーを通らない。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    data_last = row - 1
    row = note(
        ws,
        row,
        5,
        "読み方。OAuth と「API認可トークンの取得API」は Salesforce では同じ旅行の二つの名前。インターネットの業務一覧に行が無いことと、Cognito の /oauth2/token が無いことは別の事実。発行者はどちらも Cognito ユーザープール。",
        widths,
    )
    ws.auto_filter.ref = f"A4:E{data_last}"
    ws.print_area = f"A1:E{row - 1}"
    ws.page_setup.fitToHeight = 1
    ws.oddHeader.center.text = '&"Meiryo UI"&9結論'


def build_paths(ws):
    widths = [22, 28, 36, 38, 36, 32]
    apply_page(ws, "三つの入口")
    set_widths(ws, widths)
    merge_title(ws, 6, "三つの入口。合言葉の教室と、通行証の教室")
    row = section(ws, 3, 6, "比較")
    row = write_header(
        ws,
        row,
        ["項目", "イントラネットAPI", "Salesforce API", "インターネットAPI", "共通の注意", "性能試験"],
        widths,
    )
    rows = [
        [
            "たとえ",
            "校舎の中の教室。職員証の合言葉を見せる。",
            "提携バスが地下通路で入る教室。先に校門の外で通行証をもらう。",
            "町から正面玄関へ来る教室。通行証の売場は教室一覧の外。",
            "通行証と合言葉を混ぜない。",
            "インターネットの教室は今回測らない。",
        ],
        [
            "認証の種類",
            "ヘッダの API-Key。長く使う共有の秘密。",
            "OAuth 2.0 クライアントクレデンシャルズ。仕様名は API認可トークンの取得API。",
            "同じ OAuth。業務一覧にトークン取得行が無くても窓口はある。",
            "M2M は人のログイン画面を出さない。",
            "イントラは固定ヘッダ。M2M は使い回し。",
        ],
        [
            "呼び出し元が最初に出すもの",
            "業務リクエストそのもの。ヘッダに API-Key。",
            "client_id、client_secret、grant_type=client_credentials、scope。",
            "外部システムが同じトークン要求を Cognito へ出す。",
            "シークレットは業務リクエストに載せない。",
            "シークレットとトークンを jtl に残さない。",
        ],
        [
            "発行者",
            "アクセストークンの発行者はいない。合言葉は運用側が事前に登録する。",
            "Amazon Cognito ユーザープールが JWT に署名して発行する。",
            "同じ Cognito ユーザープール。アプリクライアントは分けられる。",
            "API Gateway と ECS は発行しない。検証と処理である。",
            "対象外のインターネット呼び出し元は売場へ行かせない。",
        ],
        [
            "業務リクエストの通り道",
            "呼び出し元 → イントラAPI用ALB → ECS。",
            "Salesforce → Private Connect → プライベート API Gateway → VPC Link → ALB → ECS。",
            "外部 → リージョナル API Gateway → VPC Link → GWLB / Palo Alto → ALB → ECS。",
            "トークン取得のホストと業務のホストは別。",
            "Palo Alto 経路はインターネット試験を外すと負荷から消える。",
        ],
        [
            "受け側が見るもの",
            "API-Key が登録値と一致するか。",
            "Authorization: Bearer。署名、iss、exp、token_use、client_id、スコープ。",
            "同じ Bearer 検証をリージョナル API Gateway が行う。",
            "スコープはメソッドに書く。",
            "オーソライザーキャッシュは検証回数を減らす。寿命とは別。",
        ],
        [
            "有効期間",
            "JWT の exp は無い。運用が合言葉を替えるまで。",
            "アクセストークンは発行から 60分。expires_in=3600。",
            "同じ 60分。クライアントが別ならトークンも別。",
            "設定変更後の新規発行から新しい寿命になる。既発行分の exp は残る。",
            "試験専用クライアントも本番と同じ 60分に合わせる。",
        ],
        [
            "再利用",
            "同じ API-Key を毎リクエストのヘッダへ入れる。",
            "同じアクセストークンを、前倒し時刻まで全トランザクションで使う。",
            "同じ。対象外なので試験では発生させない。",
            "Bearer は持っている人が期限まで使える。ログへ出さない。",
            "取り直しは 1本だけ。業務サンプルの外。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        6,
        "接続図の根拠は、この作業フォルダのタイムアウト整合（線4・線5・線6）、Salesforce 経路、CloudFront から ECS までの経路。インターフェース仕様書の原文はこのフォルダに無い。インターネット一覧に行が無い、という点は質問の記述を採用している。",
        widths,
    )
    ws.auto_filter.ref = "A4:F11"
    ws.print_area = f"A1:F{row - 1}"


def build_issuer(ws):
    widths = [28, 36, 42, 40, 32]
    apply_page(ws, "誰が発行するか")
    set_widths(ws, widths)
    merge_title(ws, 5, "アクセストークンの発行者。依頼、発行、検証、処理を分ける")
    row = section(ws, 3, 5, "役割")
    row = write_header(ws, row, ["役割", "イントラネット", "Salesforce", "インターネット", "M2Mでの中身"], widths)
    rows = [
        [
            "資格を持っている側",
            "社内呼び出し元が API-Key を持つ。",
            "Salesforce のアプリクライアントが ID とシークレットを持つ。",
            "外部システムのアプリクライアントが ID とシークレットを持つ。",
            "人のユーザ名は使わない。",
        ],
        [
            "依頼する",
            "業務リクエストを ALB へ送る。事前のトークン依頼は無い。",
            "Salesforce が Cognito の /oauth2/token へ POST する。これが仕様のトークン取得API。",
            "外部システムが同じ POST をする。業務一覧の行ではなく、別ホストへの依頼。",
            "grant_type=client_credentials。",
        ],
        [
            "資格を確認する",
            "受け側が API-Key を登録値と比較する。",
            "Cognito がクライアント、シークレット、許可フロー、スコープを見る。",
            "同じ Cognito の確認。",
            "シークレット不一致は invalid_client 系。ユーザの有無は見ない。",
        ],
        [
            "アクセストークンを発行する",
            "発行しない。",
            "Cognito ユーザープール。秘密鍵で JWT に署名する。",
            "同じユーザープール。GetClientToken も発行者は Cognito。",
            "応答は access_token、token_type、expires_in。",
        ],
        [
            "保管して再利用する",
            "呼び出し元が API-Key を設定として持つ。",
            "Salesforce が expires_in までキャッシュする。",
            "外部システムがキャッシュする。試験ではこの系統を撃たない。",
            "更新トークンによる延長は無い。",
        ],
        [
            "検証する",
            "イントラの受け側。API Gateway の Cognito オーソライザーはこの線にいない。",
            "プライベート API Gateway の Cognito オーソライザー。",
            "リージョナル API Gateway の Cognito オーソライザー。",
            "公開鍵は JWKS。毎リクエストで Cognito へ電話しない。",
        ],
        [
            "業務を処理する",
            "ECS 上のアプリケーション。",
            "同じアプリケーション層。トークンは作らない。",
            "同じ。検査装置（Palo Alto）は認証の発行者ではない。",
            "業務ルール（在庫など）は検証のあと。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = section(ws, row, 5, "トークン要求と業務要求は別の旅行")
    row = write_header(ws, row, ["順", "Salesforce", "インターネット", "ホスト", "試験での扱い"], widths)
    rows = [
        [
            "1. 通行証",
            "カスタムドメインの CloudFront 経由で Cognito。",
            "同じ売場。業務APIの一覧には載らないことがある。",
            "auth のドメイン。パス /oauth2/token。",
            "業務サンプルの外。使い回すなら約 55分に 1回。",
        ],
        [
            "2. 仕事",
            "Private Connect からプライベート API Gateway。",
            "インターネットからリージョナル API Gateway。",
            "業務APIのホスト。Bearer を載せる。",
            "インターネットの 2 は対象外。Salesforce の 2 は対象次第。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        5,
        "仕様書の「API認可トークンの取得API」が業務アプリのラッパーとして実装されていても、JWT の署名者は Cognito である。この作業フォルダの接続図は、呼び出し元が Cognito へ直接行く形を本線にしている。",
        widths,
    )
    ws.print_area = f"A1:E{row - 1}"


def build_reuse(ws):
    widths = [28, 22, 55, 48]
    apply_page(ws, "再利用と有効期間")
    set_widths(ws, widths)
    merge_title(ws, 4, "再利用できる。有効期間は発行から 60分")
    row = section(ws, 3, 4, "言葉の対応")
    row = write_header(ws, row, ["言葉", "値（数式）", "意味", "M2Mで起きること"], widths)
    # 値は数式で 08 を参照する
    specs = [
        ("アクセストークン有効期限", "=AccessMin", "分。コンソールの設定。既定も 60分。範囲は 5分から 1日。", "通行証の寿命。この欄が M2M に効く。"),
        ("expires_in", "=AccessMin*60", "秒。トークン応答の数値。60分なら 3600。", "呼び出し元はこの秒数を保管期限の計算に使う。"),
        ("前倒し", "=SkewSec", "秒。期限の何秒前に取り直すか。設計値 300。", "時計のずれと、オーソライザーキャッシュの端を吸収する。"),
        ("取り直し間隔", "=AccessMin*60-SkewSec", "秒。有効秒数から前倒しを引く。", "この間隔ごとに、クライアントあたり 1回発行する。"),
        ("更新トークン", '=RefreshDays&" 日"', "人の認可コードで渡る回数券の寿命。既定 30日。範囲は 60分から 10年。", "M2M の応答には入らない。30日での自動延長は無い。"),
        ("IDトークン", '=IdMin&" 分"', "人の名札の寿命。1時間は 60分と同じ長さ。", "M2M の応答には入らない。Bearer に付けるのはアクセストークン。"),
        ("認証フローセッション", '=SessionMin&" 分"', "人のチャレンジ用セッション文字列。既定 3分。範囲 3〜15分。", "通行証はここでは切れない。トークン取得の締め切りは呼び出し元の 10秒。"),
        ("オーソライザーキャッシュ", "=CacheSec", "秒。API Gateway が同じトークンの検証結果を覚える時間。既定の目安 300。範囲 0〜3600。", "トークン寿命とは別。前倒しをこれ以上にすると期限の端に頼らない。"),
    ]
    for i, (name, formula, meaning, effect) in enumerate(specs):
        bg = ZEBRA if i % 2 else WHITE
        paint(ws.cell(row, 1), name, bg=bg)
        cell = ws.cell(row, 2, formula)
        cell.font = font(10, True, INK)
        cell.alignment = align("right")
        cell.fill = fill(PAPER)
        cell.border = THIN
        if i == 0:
            cell.number_format = '0" 分"'
        elif i in (1, 2, 3, 7):
            cell.number_format = '#,##0" 秒"'
        else:
            cell.number_format = "@"
        paint(ws.cell(row, 3), meaning, bg=bg)
        paint(ws.cell(row, 4), effect, bg=bg)
        ws.row_dimensions[row].height = max(
            height_for(meaning, widths[2], cap=90),
            height_for(effect, widths[3], cap=90),
            36,
        )
        row += 1
    row = section(ws, row, 4, "再利用してよい範囲")
    row = write_header(ws, row, ["場面", "判断", "理由", "実装"], widths)
    rows = [
        [
            "同じクライアントの業務リクエスト",
            "exp の前まで何度でもよい",
            "JWT は接続に結び付かない。並行してよい。",
            "メモリの共有変数。ログには出さない。",
        ],
        [
            "期限の前倒し以降",
            "新規発行する",
            "クライアントIDとシークレットを再び送る。更新グラントは無い。",
            "取り直しは 1スレッドだけ。",
        ],
        [
            "別クライアント",
            "トークンは別",
            "スコープとシークレットが分かれる。",
            "Salesforce 用と試験用を混ぜない。",
        ],
        [
            "exp のあと",
            "署名検証は拒否する",
            "有効期間が終わる。",
            "キャッシュ TTL 中は、投入時に合格した結果が残ることがある。",
        ],
        [
            "シークレットローテーションのあと",
            "新規発行は新しいシークレットだけ",
            "既発行 JWT の署名は exp まで残る。",
            "窓を縮めたいときは寿命の設計を短くする。",
        ],
        [
            "API-Key",
            "毎リクエストで同じ値",
            "イントラネットに exp は無い。",
            "Cognito の取り直し処理を入れない。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        4,
        "数式の列はシート 08 の灰色の太枠を参照する。60分と前倒し 300秒の初期値では、取り直し間隔は 3300秒。",
        widths,
    )
    ws.print_area = f"A1:D{row - 1}"


def build_settings(ws):
    widths = [28, 18, 22, 48, 42]
    apply_page(ws, "Cognitoの各設定")
    set_widths(ws, widths)
    merge_title(ws, 5, "提示された設定。M2Mではどれが通行証に効くか")
    row = section(ws, 3, 5, "欄ごとの説明")
    row = write_header(
        ws,
        row,
        ["設定", "提示値", "M2Mの発行応答", "動作原理", "歴史と必要になった理由"],
        widths,
    )
    rows = [
        [
            "クライアントID",
            "アプリクライアントごと",
            "要求とクレームに入る",
            "ロッカー番号。公開の識別子。これだけでは発行されない。トークンの client_id と、多くの場合 sub になる。",
            "OAuth 2.0 が、人のユーザ名と別にクライアントを登録する形を選んだ。機械の台数だけ人のアカウントを増やさないため。",
        ],
        [
            "クライアントシークレット",
            "アプリクライアントごと",
            "要求で見せ、トークンには書かない",
            "金庫の鍵。Basic ヘッダか POST 本文で Cognito だけに見せる。一致したときだけ署名する。管理者は DescribeUserPoolClient で参照できる。ログへは出さない。",
            "クライアントクレデンシャルズは機密クライアント向けに作られた。鍵が漏れると、攻撃者が期限内の通行証を自分で刷れる。",
        ],
        [
            "更新トークンの有効期限",
            "30日（既定）",
            "応答に入らない",
            "人の認可コードで渡る回数券の寿命。範囲は 60分から 10年。refresh_token グラントで新しい ID とアクセスをもらう。M2M は回数券を受け取らないので、30日で延長する動作は無い。",
            "人を 60分ごとにパスワード入力へ戻さないために、2012年の OAuth 2.0 が更新トークンを分けた。機械は最初から鍵を持っているので、この回数券を使わない。",
        ],
        [
            "アクセストークンの有効期限",
            "60分（既定）",
            "これが通行証の寿命",
            "範囲は 5分から 1日。応答の expires_in は 3600秒。exp は発行時刻から 3600秒。設定を変えても、既発行トークンの exp は変わらない。",
            "Bearer は拾った人が使える。窓を無限にしないために寿命を付ける。1時間は、取り直しの回数と、漏れたあとの残存時間の釣り合いとして既定になった。",
        ],
        [
            "IDトークンの有効期限",
            "1時間（既定）",
            "応答に入らない",
            "人の名札の寿命。60分と同じ長さで、種類が違う。token_use=id。スコープを要求する API はアクセストークンを見る。",
            "OpenID Connect が、認証された人の属性を API 用のアクセス許可と分けて渡すために導入した。機械には名札の相手がいない。",
        ],
        [
            "認証フローセッションの持続期間",
            "3分（既定）",
            "トークン発行の会話に出ない",
            "InitiateAuth と RespondToAuthChallenge のセッション文字列の寿命。範囲は 3〜15分。マネージドログインの多要素は 3分、仮パスワード変更は 8分という別固定がある。M2M の POST は 1回で終わり、締め切りは呼び出し元の 10秒。",
            "パスワードのあとにもう一段の問い（コード入力）が増えたとき、問いの途中をいつまで覚えておくかを決める必要が出た。通行証の寿命とは別タイマーである。",
        ],
        [
            "トークンの取消を有効化",
            "有効（新規クライアントの既定）",
            "渡す更新トークンが無い",
            "RevokeToken と /oauth2/revoke は更新トークンを無効にする。アクセストークンを渡すと unsupported_token_type。署名と期限だけを見る検証では、取消済みでも exp まで有効なまま、と公式が書いている。",
            "人のログアウトを、30日の回数券の途中で効かせるために付いた。JWT はドアが発行者へ電話しない方式なので、即時取消と引き換えになっている。",
        ],
        [
            "ユーザ存在エラーの防止",
            "有効（コンソール既定）",
            "トークン成否の文面は変わらない",
            "ENABLED のとき、居ないユーザ名でも UserNotFound と言わない。効くのは人のサインイン、確認、パスワード回復。API 作成時の省略値は LEGACY。クライアントクレデンシャルズはユーザ名を探さない。",
            "ログイン失敗の文面が「居ない」と「パスワード違い」で分かれると、攻撃者が在籍名簿を集められた。ウェブの大規模サインイン以降、文面を揃える対策が標準になり、Cognito がスイッチにした。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        5,
        "画面に八つの欄が並んでいても、M2M の応答 JSON に載る寿命はアクセストークンの 60分だけである。人のログイン用クライアントと機械用クライアントは分け、client_credentials はそのクライアントの OAuth フローにする。ALLOW_CLIENT_TOKEN_AUTH を使う場合は、そのフローだけを付ける。",
        widths,
    )
    ws.auto_filter.ref = "A4:E11"
    ws.print_area = f"A1:E{row - 1}"


def build_motion(ws):
    widths = [16, 42, 42, 42, 28]
    apply_page(ws, "動作イメージ")
    set_widths(ws, widths)
    merge_title(ws, 5, "動作イメージ。0秒、55分、60分")
    row = section(ws, 3, 5, "Salesforce または外部システムがトークンを使うとき")
    row = write_header(ws, row, ["時刻", "呼び出し元", "Cognito", "API Gateway", "ECS"], widths)
    rows = [
        [
            "0.0秒",
            "クライアントIDとシークレットを /oauth2/token へ送る。",
            "アプリクライアント、シークレット、フロー、スコープを見る。",
            "この旅行にはまだ出てこない。",
            "まだ何も受けない。",
        ],
        [
            "約 0.5秒",
            "access_token と expires_in=3600 をメモリへ置く。",
            "JWT に署名して返す。IDトークンと更新トークンは付けない。",
            "まだ見ていない。",
            "まだ受けない。",
        ],
        [
            "その後の各業務",
            "Authorization: Bearer に同じ文字列を付ける。売場へは戻らない。",
            "電話は来ない。ドアは公開鍵で見る。",
            "WAF のあと、署名、iss、exp、token_use、スコープを見る。合格結果はキャッシュすることがある。",
            "業務だけを処理する。トークンは作らない。",
        ],
        [
            "3300秒",
            "前倒し。1本だけが新しいトークンを取り、共有する。",
            "新しい JWT を発行する。古い JWT の exp はまだ残る。",
            "古いトークンも、あと 300秒は署名上有効。",
            "業務は途切れさせない。",
        ],
        [
            "3600秒",
            "古いトークンを Bearer に付けるのをやめる。",
            "古いトークンの新規発行ではない。exp を書き換えない。",
            "期限切れとして拒否する。キャッシュが残っていれば TTL まで許可が残ることがある。",
            "拒否されたリクエストは届かない。",
        ],
        [
            "10秒で諦める線",
            "トークン取得が 10秒を超えたら、その取得だけを失敗にする。",
            "普段の応答は 0.5秒目標。10秒は上限であり、普段から使い切らない。",
            "業務側の上限は別。API Gateway 統合は 29秒。",
            "トークン失敗のリクエストは業務へ進まない。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = section(ws, row, 5, "イントラネットは別の時計")
    row = write_header(ws, row, ["時刻", "呼び出し元", "Cognito", "入口", "ECS"], widths)
    rows = [
        [
            "毎リクエスト",
            "ヘッダに API-Key を入れて ALB へ送る。",
            "この線には登場しない。",
            "イントラAPI用ALB。公開 API Gateway は無い。",
            "合言葉の一致を見て業務を処理する。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        5,
        "3分の認証フローセッションは、この表のどこにも通行証の締め切りとして出ない。人の追加質問に答える時間である。",
        widths,
    )
    ws.print_area = f"A1:E{row - 1}"


def build_history(ws):
    widths = [18, 40, 52, 40]
    apply_page(ws, "歴史")
    set_widths(ws, widths)
    merge_title(ws, 4, "なぜこの三種類が並んだか")
    row = section(ws, 3, 4, "年表")
    row = write_header(ws, row, ["時期", "何が生まれたか", "必要になった経緯", "今回の三経路との関係"], widths)
    rows = [
        [
            "1961年",
            "CTSS でパスワードが使われ始めた。",
            "部屋に入れる人だけ、では時間共有の利用者が分からなくなった。",
            "イントラネット API-Key の先祖。知っている者だけが通る合言葉。",
        ],
        [
            "2006〜2010年",
            "OAuth の議論と、RFC 5849（OAuth 1.0）。",
            "別のアプリへ自分のパスワードを預けないと委任できなかった。署名の手順は重かった。",
            "パスワードそのものを業務APIへ貼らない、という方向が始まった。",
        ],
        [
            "2012年10月",
            "RFC 6749 OAuth 2.0 と RFC 6750 Bearer。クライアントクレデンシャルズは 4.4。",
            "人の委任と、機械自身の呼び出しを分け、ドアで見せるものを期限つきトークンにした。",
            "Salesforce とインターネットの認証方式が、この節である。",
        ],
        [
            "2015年5月",
            "RFC 7519 JWT。",
            "ドアが発行者へ毎回帰電すると、売場が全APIの障害点になる。",
            "60分の再利用が性能上正しい理由。取消が即座には効かない理由でもある。",
        ],
        [
            "2014〜2016年",
            "Cognito の登場と、ユーザープール。2015年7月に API Gateway。",
            "アプリがパスワード台帳を自前だけで持たず、扉をアプリの前に置きたくなった。",
            "発行者（Cognito）と検証者（API Gateway）を分けられるようになった。",
        ],
        [
            "2020年頃",
            "ユーザープールでクライアントクレデンシャルズ。",
            "人の発行者と機械の発行者を二つ運用する負担をまとめたかった。",
            "Salesforce とインターネットが同じユーザープールを向ける。クライアントは分けられる。",
        ],
        [
            "2024-05-09",
            "M2M トークン要求の料金が明示された。",
            "機械の認可がユーザ認証のおまけでは数えられなくなった。",
            "毎回取得すると、性能の前に回数と料金が膨らむ。",
        ],
        [
            "2024-11-22",
            "Lite / Essentials / Plus。新規の既定は Essentials。",
            "人のログイン機能と、追加のセキュリティをプランで分けた。",
            "M2M のクレーム調整は Essentials 以上の機能。料金は別建て。",
        ],
        [
            "2025-03-03 と 04-30",
            "M2M アクセストークンの調整と、クライアントメタデータ。",
            "スコープだけでは細かい許可をトークンへ書けなかった。",
            "クレームを足すとヘッダが伸びる。プライベート API は合計 8000バイト。",
        ],
        [
            "2025-11-03",
            "アプリクライアント個数の月額が外れた。",
            "窓口の登録数ではなく、成功した発行回数を数える形に寄った。",
            "試験専用クライアントを作ること自体の月額は、現行説明では加算されない。",
        ],
        [
            "2026-08-31",
            "GetClientToken。",
            "ドメイン、証明書、DNS 無しで M2M トークンが欲しくなった。",
            "OAuth の本線は /oauth2/token のまま。別窓口であり、業務一覧の行ではない。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        4,
        "イントラネットだけが合言葉のままなのは、壁の内側で、毎リクエストの JWT 発行より固定ヘッダの方が軽いからである。外側の Salesforce とインターネットは、漏れた通行証が 60分で終わる分離を選んでいる。",
        widths,
    )
    ws.auto_filter.ref = "A4:D14"
    ws.print_area = f"A1:D{row - 1}"


def build_perf(ws):
    widths = [38, 28, 55, 36]
    apply_page(ws, "性能の計算")
    set_widths(ws, widths)
    merge_title(ws, 4, "性能試験。インターネットを外したときの発行回数")
    row = section(ws, 3, 4, "計算結果。入力はシート 08")
    row = write_header(ws, row, ["項目", "値", "式の意味", "読み方"], widths)
    calcs = [
        ("対象の M2M クライアント数", "=SfClients+NetClients", "Salesforce 用 + インターネット用。初期値は 1+0。", "インターネットを外すと、ここへ 0 を足す。"),
        ("アクセストークンの秒数", "=AccessMin*60", "60分 → 3600秒。", "expires_in と一致させる。"),
        ("取り直し間隔（秒）", "=AccessMin*60-SkewSec", "寿命から前倒しを引く。", "0以下なら前倒しが寿命以上。下の判定を見る。"),
        ("間隔は正か", '=IF(AccessMin*60-SkewSec>0,"正","前倒しを短くする")', "取り直しが成立する条件。", "初期値 3300秒で正。"),
        ("使い回したときのトークン成功回数", "=IF(AccessMin*60-SkewSec<=0,NA(),(SfClients+NetClients)*Hours*3600/(AccessMin*60-SkewSec))", "クライアント × 試験時間 × 3600 ÷ 間隔。", "初期値は約 1.09 回。業務の回数とは桁が違う。"),
        ("毎回取得したときのトークン成功回数", "=BizRps*Hours*3600", "対象の M2M 業務 RPS × 秒。イントラの API-Key 呼び出しは含めない。", "業務のたびに売場へ行くと、この回数になる。"),
        ("毎回取得 ÷ 使い回し", "=IF(OR(AccessMin*60-SkewSec<=0,SfClients+NetClients=0),NA(),(BizRps*Hours*3600)/((SfClients+NetClients)*Hours*3600/(AccessMin*60-SkewSec)))", "発行回数が何倍に膨らむか。", "初期値は 16万倍の規模。"),
        ("使い回しの例示料金（USD）", "=IF(AccessMin*60-SkewSec<=0,NA(),(SfClients+NetClients)*Hours*3600/(AccessMin*60-SkewSec)*UnitPrice)", "成功回数 × 単価。単価の初期値は米国東部の例。", "東京の単価は料金ページで UnitPrice のセルを差し替える。"),
        ("毎回取得の例示料金（USD）", "=BizRps*Hours*3600*UnitPrice", "同じ単価での比較。", "試験のたびにこの額を発行だけで使う、という例。"),
        ("イントラネットのトークン回数", "=0", "API-Key 経路は Cognito を呼ばない。", "この 0 は固定。インターネット対象外とは独立。"),
        ("キャッシュの秒", "=CacheSec", "オーソライザー結果の TTL。", "トークン取り直し間隔には足さない。検証回数の話。"),
    ]
    for i, (name, formula, meaning, how) in enumerate(calcs):
        bg = ZEBRA if i % 2 else WHITE
        paint(ws.cell(row, 1), name, bg=bg, bold=True)
        cell = ws.cell(row, 2, formula)
        cell.font = font(11, True, INK)
        cell.alignment = align("right")
        cell.fill = fill(PAPER)
        cell.border = THIN
        if i == 4:
            cell.number_format = "#,##0.00"
        elif i in (5, 6):
            cell.number_format = "#,##0.00"
        elif i in (7, 8):
            cell.number_format = '"$"#,##0.000000'
        elif i in (0, 1, 2, 10):
            cell.number_format = "#,##0"
        paint(ws.cell(row, 3), meaning, bg=bg)
        paint(ws.cell(row, 4), how, bg=bg)
        ws.row_dimensions[row].height = max(height_for(meaning, widths[2], cap=72), height_for(how, widths[3], cap=72), 32)
        row += 1
    row = section(ws, row, 4, "大勢が変わらない理由")
    row = write_header(ws, row, ["観点", "判断", "根拠", "残る注意"], widths)
    rows = [
        [
            "インターネットを対象外にする",
            "試験全体の大勢は変わらない",
            "公開 API Gateway、WAF、Palo Alto、インターネット呼び出し元のトークン要求が負荷から外れる。",
            "Salesforce を対象に残すなら、そのクライアントの使い回しだけは残る。",
        ],
        [
            "一覧にトークン取得が無い",
            "隠れた大量通信にはならない",
            "窓口は業務メソッドではなく Cognito の別ホスト。対象外の呼び出し元だけが行く。",
            "対象内のシナリオへ、毎リクエストのトークン取得を書き足さない。",
        ],
        [
            "イントラネット",
            "Cognito 設定はサンプル時間を変えない",
            "認証は API-Key のヘッダ比較。60分も 3分もこの線を通らない。",
            "キーの実値を結果へ残さない。",
        ],
        [
            "使い回しを忘れた場合",
            "大勢が変わるのはここ",
            "発行回数が業務 RPS と同じになり、クォータと料金と待ち時間が先に頭打ちになる。",
            "トークン時間は業務トランザクションの外へ置く。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        4,
        "BizRps の初期値 50 は例である。実測の同時実行数ではない。料金は成功したトークン応答の例示で、M2M に無料枠は無い。アプリクライアントの個数への追加料金は、確認日の料金説明には無い。",
        widths,
    )
    ws.print_area = f"A1:D{row - 1}"
    # 間隔が不正なとき、判定セルの隣を強調するルールは値の側だけでよい
    ws.conditional_formatting.add(
        "B8",
        FormulaRule(formula=['$B$8="前倒しを短くする"'], fill=fill("D9D9D9"), font=font(11, True, INK)),
    )


def build_inputs(ws):
    widths = [34, 28, 22, 62]
    apply_page(ws, "調整つまみ")
    set_widths(ws, widths)
    merge_title(ws, 4, "調整つまみ。灰色の太枠だけが入力")
    row = section(ws, 3, 4, "入力")
    row = write_header(ws, row, ["説明", "名前", "入力値", "単位と根拠"], widths)
    inputs = [
        ("AccessMin", "アクセストークンの有効期限", 60, "分。提示設定。M2M の通行証の寿命。範囲の目安は 5〜1440。"),
        ("IdMin", "IDトークンの有効期限", 60, "分。提示の 1時間。M2M の応答には使われない。比較用に残す。"),
        ("RefreshDays", "更新トークンの有効期限", 30, "日。提示設定。M2M の応答には使われない。既定値。"),
        ("SessionMin", "認証フローセッション", 3, "分。提示設定。人のチャレンジ用。範囲は 3〜15。"),
        ("SkewSec", "取り直しの前倒し", 300, "秒。接続設計の値。オーソライザーキャッシュの既定目安と同じ 300。"),
        ("SfClients", "試験に含める Salesforce 用クライアント数", 1, "個。対象に Salesforce が残る例。0 にするとトークン回数は 0。"),
        ("NetClients", "試験に含めるインターネット用クライアント数", 0, "個。対象外なので初期値 0。1 以上にすると発行回数へ足される。"),
        ("BizRps", "対象 M2M の業務リクエスト毎秒（例）", 50, "回/秒。実測値ではない。イントラネットの API-Key 呼び出しは含めない。"),
        ("Hours", "試験時間", 1, "時間。発行回数の見積もり窓。"),
        ("UnitPrice", "M2M 成功 1回の例示単価", 0.00225, "USD。料金ページの米国東部（バージニア北部）の例。東京は表を見て差し替える。"),
        ("CacheSec", "オーソライザー結果キャッシュ", 300, "秒。API Gateway の既定の目安。設定範囲は 0〜3600。実環境の値で差し替える。"),
    ]
    first = row
    for i, (name, label, value, unit) in enumerate(inputs):
        bg = ZEBRA if i % 2 else WHITE
        paint(ws.cell(row, 1), label, bg=bg, bold=True)
        paint(ws.cell(row, 2), name, bg=bg)
        cell = ws.cell(row, 3, value)
        cell.font = font(12, True, INK)
        cell.alignment = align("right")
        cell.fill = fill(INPUT_BG)
        cell.border = THICK
        cell.protection = Protection(locked=False)
        if name == "UnitPrice":
            cell.number_format = "0.000000"
        elif name == "Hours":
            cell.number_format = "0.00"
        else:
            cell.number_format = "#,##0"
        paint(ws.cell(row, 4), unit, bg=bg)
        ws.row_dimensions[row].height = height_for(unit, widths[3], minimum=28, cap=72)
        row += 1
    last = row - 1
    dv_positive = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=False)
    dv_positive.error = "0より大きい数を入れる"
    dv_positive.errorTitle = "入力"
    dv_positive.add(f"C{first}:C{first + 4}")
    dv_positive.add(f"C{first + 7}:C{first + 8}")
    ws.add_data_validation(dv_positive)
    dv_nonneg = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=False)
    dv_nonneg.error = "0以上の数を入れる"
    dv_nonneg.errorTitle = "入力"
    dv_nonneg.add(f"C{first + 5}:C{first + 6}")
    dv_nonneg.add(f"C{first + 9}:C{first + 10}")
    ws.add_data_validation(dv_nonneg)
    row = section(ws, row, 4, "入力からその場で読める派生値")
    row = write_header(ws, row, ["派生", "式", "値", "意味"], widths)
    derived = [
        ("アクセストークン秒", "=AccessMin*60", "expires_in に相当する秒。"),
        ("IDトークン秒", "=IdMin*60", "人の名札の秒。M2M 応答には出ない。"),
        ("更新トークン秒", "=RefreshDays*24*60*60", "人の回数券の秒。M2M 応答には出ない。"),
        ("セッション秒", "=SessionMin*60", "人のチャレンジの秒。通行証の寿命ではない。"),
        ("取り直し間隔秒", "=AccessMin*60-SkewSec", "この値が正のとき、使い回しの周期になる。"),
        ("対象クライアント", "=SfClients+NetClients", "インターネット初期値 0 を含む合計。"),
        ("1時間あたりの使い回し回数", "=IF(AccessMin*60-SkewSec<=0,NA(),(SfClients+NetClients)*3600/(AccessMin*60-SkewSec))", "試験時間を 1時間に固定した参考。"),
    ]
    for i, (name, formula, meaning) in enumerate(derived):
        bg = ZEBRA if i % 2 else WHITE
        paint(ws.cell(row, 1), name, bg=bg, bold=True)
        paint(ws.cell(row, 2), "数式", bg=bg)
        cell = ws.cell(row, 3, formula)
        cell.font = font(12, True, INK)
        cell.alignment = align("right")
        cell.fill = fill(PAPER)
        cell.border = THIN
        cell.number_format = "#,##0.00"
        paint(ws.cell(row, 4), meaning, bg=bg)
        ws.row_dimensions[row].height = 28
        row += 1
    row = note(
        ws,
        row,
        4,
        "C列の灰色太枠が入力。同じ列の下段と、シート 07 の数値は数式。名前付き範囲 AccessMin などは C5 から C15。NetClients を 0 のままにするのが、インターネット対象外の表現。",
        widths,
    )
    ws.print_area = f"A1:D{row - 1}"
    ws.auto_filter.ref = f"A4:D{last}"


def build_sources(ws):
    widths = [28, 42, 55, 30]
    apply_page(ws, "出典")
    set_widths(ws, widths)
    merge_title(ws, 4, "出典と、未確認の境界")
    row = section(ws, 3, 4, "2026-10-14 に確認した公式")
    row = write_header(ws, row, ["系統", "文書", "今回使った事実", "境界"], widths)
    rows = [
        [
            "Cognito",
            "The token endpoint",
            "クライアントクレデンシャルズはアクセストークンだけ。更新トークンは認可コードのとき。",
            "確認日の英語版ドキュメント。",
        ],
        [
            "Cognito",
            "Refresh tokens / token revocation / user existence errors",
            "更新は 30日既定。取消は更新トークン。署名検証だけでは exp まで残る。ユーザ存在の防止は人の API。",
            "M2M へ読み替えた適用範囲は本文の判断。",
        ],
        [
            "Cognito",
            "認証フローセッション、GetClientToken",
            "セッション既定 3分、範囲 3〜15。GetClientToken は 2026-08-31 案内以降の別窓口。",
            "本線は仕様の OAuth、つまり /oauth2/token。",
        ],
        [
            "料金",
            "Amazon Cognito pricing の M2M 節",
            "成功応答への課金。個数への追加料金の記載は無い。無料枠は M2M に無い。米国東部の例 0.00225 USD。",
            "東京単価はページの表が正。セルで差し替える。",
        ],
        [
            "AWS Security Blog",
            "2025-01-13 M2M の監視と保護",
            "有効なあいだキャッシュして再利用する。",
            "実装指針。クォータ数値の代替にはしない。",
        ],
        [
            "この作業フォルダ",
            "タイムアウト整合、Salesforce 経路、CloudFront から ECS",
            "線4は API-Key に相当する ALB 直結。線5と線6は Cognito でトークンを取ってから業務。寿命 3600秒、前倒し 300秒。",
            "仕様書の原文はフォルダに無い。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = section(ws, row, 4, "未確認")
    row = write_header(ws, row, ["項目", "状態", "結論への影響", "確認する場所"], widths)
    rows = [
        [
            "インターフェース仕様の原文",
            "未収録",
            "インターネット一覧に行が無い、という質問の記述を採用した。発行者が Cognito であることは接続図と一致する。",
            "仕様書の認証章と URL。",
        ],
        [
            "機能プランと実 TTL と東京単価",
            "アカウントは開いていない",
            "60分の読みと、対象外でも大勢が変わらない、という構造はプランが違っても崩れない。金額の桁は単価セルに依存する。",
            "コンソールと料金表、Service Quotas。",
        ],
        [
            "API-Key の比較主体",
            "ALB の先がアプリであることまで",
            "比較がアプリでも前段の社内装置でも、Cognito が発行しないことは変わらない。",
            "受信フィルタの実装。",
        ],
    ]
    row = write_rows(ws, row, rows, widths)
    row = note(
        ws,
        row,
        4,
        "Markdown と同じフォルダにある。Excel の数値を変えたあとは、シート 07 を再計算してから読む。",
        widths,
    )
    ws.print_area = f"A1:D{row - 1}"


if __name__ == "__main__":
    print(build())
