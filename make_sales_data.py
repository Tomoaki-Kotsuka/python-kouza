"""講座用の元データ sales_data.xlsx を作るスクリプト。

使い方:  python make_sales_data.py
乱数のタネ(シード)を固定しているので、何度実行しても同じデータになります。
ノートブック(python_excel_kouza.ipynb)の「保険セル」にも、
下の「ここから」〜「ここまで」と同じコードが入っています。
"""

# ==== ここから ====
# 乱数(サイコロ)を使うための道具を読み込む
import random
# その月が何日まであるかを調べる道具を読み込む
import calendar
# 表を扱う道具 pandas を pd という名前で読み込む
import pandas as pd


# 売上データの Excel ファイルを作る手順をまとめて、make_sales_data という名前を付ける
def make_sales_data(path="sales_data.xlsx", seed=42):
    # 乱数のタネを固定する(毎回まったく同じデータになる)
    rng = random.Random(seed)
    # 地区ごとの設定: (月の売上の目安[万円], 月の明細件数の目安, 担当者)
    areas = {
        "札幌": (150, 4, ["佐藤 健一", "高橋 美咲"]),
        "仙台": (110, 4, ["伊藤 大輔", "渡辺 由美", "菅原 拓也"]),
        "新潟": (100, 3, ["小林 誠", "五十嵐 恵"]),
        "金沢": (80, 3, ["中村 翔太", "北川 綾香"]),
        "長野": (80, 3, ["宮沢 浩二", "丸山 千尋"]),
        "宇都宮": (80, 3, ["鈴木 直樹", "阿久津 真理"]),
        "さいたま": (160, 4, ["新井 雄太", "関根 沙織", "金子 亮"]),
        "千葉": (140, 4, ["石井 達也", "川島 奈々"]),
        "東京": (420, 8, ["田中 一郎", "山本 さくら", "松本 健太"]),
        "横浜": (240, 6, ["加藤 隼人", "小川 麻衣", "原田 修"]),
        "静岡": (120, 3, ["望月 剛", "杉山 香織"]),
        "名古屋": (260, 6, ["水野 博之", "服部 愛", "近藤 悠"]),
        "京都": (150, 4, ["西村 和也", "奥田 美穂"]),
        "大阪": (330, 7, ["吉田 康平", "森本 絵里", "岡本 大地"]),
        "神戸": (150, 4, ["藤原 智史", "井上 彩"]),
        "岡山": (100, 3, ["三宅 俊介", "難波 裕子"]),
        "広島": (140, 4, ["村上 洋平", "藤井 友香", "山根 徹"]),
        "高松": (80, 3, ["大西 正樹", "真鍋 理恵"]),
        "福岡": (130, 5, ["古賀 竜也", "松尾 美紀", "江口 航"]),
        "那覇": (60, 3, ["比嘉 優", "金城 あかね"]),
    }
    # 商品の一覧: (商品名, カテゴリ, 単価)
    products = [
        ("コピー用紙A4(5000枚)", "文房具", 4200),
        ("ボールペン(10本セット)", "文房具", 1200),
        ("フラットファイル(50冊)", "文房具", 2800),
        ("オフィスチェア", "オフィス家具", 24800),
        ("昇降デスク", "オフィス家具", 49800),
        ("収納キャビネット", "オフィス家具", 32000),
        ("ワイヤレスマウス", "PC周辺機器", 3200),
        ("外付けキーボード", "PC周辺機器", 5800),
        ("27インチモニター", "PC周辺機器", 36000),
        ("Webカメラ", "PC周辺機器", 7400),
        ("シュレッダー", "オフィス機器", 28000),
        ("ラベルプリンター", "オフィス機器", 15800),
    ]
    # 月ごとの営業日数(2023年4月〜2026年4月の37か月ぶん。平日の数から祝日を引いたもの。最後の1つは「来月」)
    workdays = [20, 20, 22, 20, 22, 20, 21, 20, 21, 21, 19, 20, 21, 21, 20, 22, 21, 19, 22, 20, 22, 21, 18, 20, 21, 20, 21, 22, 20, 20, 22, 18, 23, 20, 18, 21, 21]
    # 来月(2026年4月)にキャンペーンを予定している地区
    next_campaign = ["仙台", "金沢", "福岡", "那覇"]
    # 基本の列の並びを決める
    columns = ["売上日", "地区", "担当者", "商品名", "カテゴリ", "単価", "数量", "売上金額"]
    # シートごとの列の並び(2・3シート目はわざとバラバラにする)
    sheet_columns = {
        2023: columns,
        2024: ["担当者", "地区", "売上日", "カテゴリ", "商品名", "数量", "単価", "売上金額"],
        2025: ["売上金額", "商品名", "カテゴリ", "売上日", "数量", "単価", "地区", "担当者"],
    }
    # 年度ごとの明細を入れる箱を用意する
    rows = {2023: [], 2024: [], 2025: []}
    # 月次データ(年月・地区・営業日数・キャンペーン)を入れる箱を用意する
    cond_rows = []
    # 2023年4月から数えて何か月目かを数える数字(最初は 0)
    t = 0
    # 年度を 2023 → 2024 → 2025 の順にくり返す
    for fy in [2023, 2024, 2025]:
        # 月を 4月 → 翌年3月 の順にくり返す
        for month in [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]:
            # カレンダー上の年を、まず年度と同じにしておく
            year = fy
            # もし 1〜3月なら
            if month <= 3:
                # カレンダー上は翌年になる
                year = fy + 1
            # その月が何日まであるかを調べる
            last_day = calendar.monthrange(year, month)[1]
            # 「2023-04」のような年月の文字を作る
            ym = str(year) + "-" + str(month).zfill(2)
            # その月の営業日数を取り出す
            days = workdays[t]
            # 地区を1つずつ取り出してくり返す
            for area in areas:
                # その地区の設定を取り出す
                base, n_base, staff = areas[area]
                # キャンペーンは、まず「なし(0)」にしておく
                campaign = 0
                # 3割くらいの確率で
                if rng.random() < 0.3:
                    # キャンペーン「あり(1)」にする
                    campaign = 1
                # 条件による売上の倍率は、まず 1.00 倍(ふつうの月)にしておく
                effect = 1.00
                # 営業日数が多い(21日以上)だけなら
                if days >= 21 and campaign == 0:
                    # 少しだけ増える
                    effect = 1.08
                # キャンペーンありだけなら
                if days < 21 and campaign == 1:
                    # 少しだけ増える
                    effect = 1.10
                # 営業日数が多い月にキャンペーンが重なると
                if days >= 21 and campaign == 1:
                    # 売上がはっきり跳ねる
                    effect = 1.50
                # 月次データに1行追加する
                cond_rows.append([ym, area, days, campaign])
                # その月の売上の目安 = 規模 × 条件による倍率 × 少しのゆらぎ
                target = base * 10000 * effect * rng.gauss(1, 0.05)
                # その月の明細件数を決める(目安から -1〜+1 件ゆらす)
                n = n_base + rng.choice([-1, 0, 0, 1])
                # 月の売上を明細ごとに分ける割合を入れる箱を用意する
                weights = []
                # その月に売れた商品を入れる箱を用意する
                items = []
                # 明細の件数ぶんくり返す
                for i in range(n):
                    # 割合を少しゆらして箱に追加する
                    weights.append(rng.uniform(0.6, 1.4))
                    # 商品を1つ選んで箱に追加する
                    items.append(rng.choice(products))
                # 商品を単価の高い順に並べる(端数を安い商品で調整するため)
                items.sort(key=lambda item: item[2], reverse=True)
                # 前の明細で合わせきれなかった端数(最初は 0)
                carry = 0
                # 明細を1件ずつ作る
                for i in range(n):
                    # 商品名・カテゴリ・単価を取り出す
                    name, category, price = items[i]
                    # この明細で売りたい金額(前の端数も足す)
                    amount = target * weights[i] / sum(weights) + carry
                    # 目安の金額になるように数量を決める(最低1個)
                    qty = max(1, round(amount / price))
                    # 合わせきれなかった端数を次の明細に回す
                    carry = amount - price * qty
                    # 売上日を決める
                    day = pd.Timestamp(year, month, rng.randint(1, last_day))
                    # 担当者を選ぶ
                    person = rng.choice(staff)
                    # 1行分のデータを箱に追加する(売上金額 = 単価 × 数量)
                    rows[fy].append([day, area, person, name, category, price, qty, price * qty])
            # 1か月進める
            t = t + 1
    # 来月(2026年4月)の予定を、地区ごとに月次データへ追加する
    for area in areas:
        # キャンペーン予定は、まず「なし(0)」にしておく
        campaign = 0
        # 予定している地区なら
        if area in next_campaign:
            # 「あり(1)」にする
            campaign = 1
        # 月次データに1行追加する(売上はまだ無い)
        cond_rows.append(["2026-04", area, workdays[36], campaign])
    # Excel ファイルを書き込み用に開く
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # 年度ごとに1シートずつ書き込む
        for fy in [2023, 2024, 2025]:
            # 箱の中身を表にする
            df = pd.DataFrame(rows[fy], columns=columns)
            # 売上日の古い順に並べ替える
            df = df.sort_values("売上日", kind="stable")
            # 列をそのシート用の並びにする
            df = df[sheet_columns[fy]]
            # シート名を「2023年度」のように決める
            sheet_name = str(fy) + "年度"
            # シートに書き込む
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # A〜H 列を1つずつ取り出してくり返す
            for letter in "ABCDEFGH":
                # 列の幅を広げて見やすくする
                writer.sheets[sheet_name].column_dimensions[letter].width = 16
            # 売上日の列が左から何番目(A〜H)かを調べる
            date_letter = "ABCDEFGH"[sheet_columns[fy].index("売上日")]
            # 売上日の列のセルを、見出しをのぞいて1つずつ取り出す
            for cell in writer.sheets[sheet_name][date_letter][1:]:
                # 日付を 2023/04/01 の形で表示する設定にする
                cell.number_format = "yyyy/mm/dd"
            # 商品名の列が左から何番目(A〜H)かを調べる
            name_letter = "ABCDEFGH"[sheet_columns[fy].index("商品名")]
            # 商品名の列だけさらに広げる
            writer.sheets[sheet_name].column_dimensions[name_letter].width = 26
        # 月次データの箱の中身を表にする
        df_cond = pd.DataFrame(cond_rows, columns=["年月", "地区", "営業日数", "キャンペーン"])
        # 4枚目のシート「月次データ」に書き込む
        df_cond.to_excel(writer, sheet_name="月次データ", index=False)
        # A〜D 列を1つずつ取り出してくり返す
        for letter in "ABCD":
            # 列の幅を広げて見やすくする
            writer.sheets["月次データ"].column_dimensions[letter].width = 14
# ==== ここまで ====


if __name__ == "__main__":
    make_sales_data("sales_data.xlsx")
    print("sales_data.xlsx を作成しました")
