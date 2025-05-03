import discord
from discord.ext import commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json
import pytz #20250503 R.TSURUTA(日本時間実装のため)
# 設定ファイルを読み込む
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# Discord Botの設定
discord_token = config["discord_token"]
google_creds_file = config["google_creds_file"]
spreadsheet_name = config["spreadsheet_name"]

# Discordのインテント設定
intents = discord.Intents.default()
intents.voice_states = True
intents.members = True
intents.message_content = True

# Botのインスタンス作成
bot = commands.Bot(command_prefix="!", intents=intents)

# 日本標準時 (JST) のタイムゾーンを指定　20250503　R.TSURUTA
jst = pytz.timezone('Asia/Tokyo')

# Google Sheets認証
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(google_creds_file, scope)
client = gspread.authorize(creds)
sheet = client.open(spreadsheet_name).sheet1

# VC参加者のトラッキング辞書
vc_tracking = {}
vc_active = {}  # VCに参加しているユーザーの活動状態を追跡

# VC終了時間および接続時間の更新処理
def update_vc_end_time_and_duration(member_id: str, end_time: str):
    rows = sheet.get_all_values()  # 全行（ヘッダー含む）を取得
    target_row_f = None  # F列の空白がある行
    target_row_g = None  # G列の空白がある行

    # 下から上に逆順で行を探していきます
    for i in range(len(rows) - 1, 0, -1):  # 1行目はヘッダーなので飛ばす
        row = rows[i]
        
        # IDが一致し、F列が空白の行を探す *かつ、D列がVCの場合も追加
        if row[0] == member_id and row[3] == "VC" and row[5] == "" and target_row_f is None:
            target_row_f = i + 1  # gspreadは1始まりのインデックス
            
        # IDが一致し、G列が空白の行を探す かつ、D列がVCの場合も追加
        if row[0] == member_id and row[6] == "" and target_row_g is None:
            target_row_g = i + 1  # gspreadは1始まりのインデックス
            
        # 両方の条件を満たす行を探し終えたら、ループを抜ける
        if target_row_f and target_row_g:
            break

    # F列が空白の行に終了時間を更新
    if target_row_f:
        sheet.update_cell(target_row_f, 6, end_time)  # F列（終了時間）を更新
        print(f"Updated end time for user {member_id} in row {target_row_f}.")
    else:
        print(f"No matching row found for user {member_id} with empty end time (F column).")

    # G列が空白の行に接続時間を更新
    if target_row_g:
        sheet.update_cell(target_row_g, 7, f'=INT((F{target_row_g}-E{target_row_g})*86400)')  # G列（接続時間）を更新
        print(f"Updated duration for user {member_id} in row {target_row_g}.")
    else:
        print(f"No matching row found for user {member_id} with empty duration (G column).")

# VC参加時間を計算して更新(未使用？)
def calculate_vc_duration(sheet):
    rows = sheet.get_all_values()[1:]  # ヘッダーを除いたすべての行を取得
    
    for row in rows:
        try:
            # E列（VC接続開始時間）とF列（VC接続終了時間）の値を取得
            start_time_str = row[4]  # E列
            end_time_str = row[5]    # F列

            if start_time_str and end_time_str:
                # 文字列をdatetime型に変換
                start_time = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
                end_time = datetime.strptime(end_time_str, "%Y-%m-%d %H:%M:%S")

                # 接続時間（秒）を計算
                vc_duration = (end_time - start_time).total_seconds()

                # G列にVC接続時間を更新
                row_number = rows.index(row) + 2  # G列を更新するための行番号
                sheet.update_cell(row_number, 7, vc_duration)  # G列（VC接続時間）に更新

                print(f"VC duration for {row[1]}: {vc_duration} seconds")
        except Exception as e:
            print(f"Error calculating VC duration for row {row}: {e}")

@bot.event
async def on_voice_state_update(member, before, after):
    #入室時
    if after.channel is not None:
        channel = after.channel
        member_count = len(channel.members)
        print(f"チャンネル名:{channel.name},メンバー数:{member_count}")
        if member_count >= 2:  # VCに2人以上がいる場合
            # 日本時間を追加し、表示形式をデフォルトに変更 20250503 R.TSURUTA
            #now = datetime.now(jst).strftime("%Y-%m-%d %H:%M:%S")
            now = datetime.now(jst).strftime("%Y/%m/%d %H:%M:%S")

            # 新たに参加したユーザーを記録
            if member.id not in vc_tracking:
                vc_tracking[member.id] = now  # 参加したユーザーの開始時刻を記録
                # 新しい行に記録を追加
                sheet.append_row([ 
                    str(member.id),  # ユーザーID (A列)
                    member.name,  # ユーザー名 (B列)
                    now,  # ログ記載時間 (C列)
                    "VC",  # 種別 (D列)
                    now,  # VC接続開始時間 (E列)
                    ""  # VC接続終了時間 (F列, 退出時に記録)
                ])
                print(f"{member.name} joined {channel.name} at {now}")

            # VCに参加しているすべてのユーザーの記録を開始（新規参加ユーザーを除く）
            for m in channel.members:
                if m.id not in vc_tracking:
                    vc_tracking[m.id] = now  # 既に参加しているユーザーの記録も開始

                    # 新しい行を追加してそのユーザーの記録を追加
                    sheet.append_row([
                        str(m.id),
                        m.name,
                        now,
                        "VC",
                        now,
                        ""
                    ])
                    print(f"{m.name}'s record started as well.")

    # 退出時
    if before.channel is not None and (after.channel is None or before.channel != after.channel):
        channel = before.channel
        if member.id in vc_tracking:
            start_time = vc_tracking.pop(member.id)  # 退出したユーザーの開始時刻を取得
            # 日本時間を追加し、表示形式をデフォルトに変更 20250503 R.TSURUTA
            #end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
            end_time = datetime.now(jst).strftime("%Y/%m/%d %H:%M:%S")

            # 開始時間と終了時間をdatetime型に変換
            #start_time_obj = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S") #時間の形式がデフォルトになった影響 20250503 R.TSURUTA
            #end_time_obj = datetime.strptime(end_time, "%Y-%m-%d %H:%M:%S") #時間の形式がデフォルトになった影響 20250503 R.TSURUTA

            # 接続時間（秒）を計算(未使用のため削除) 20250503 R.TSURUTA
            #vc_duration = (end_time_obj - start_time_obj).total_seconds()

            # VC接続終了時間と接続時間を更新
            update_vc_end_time_and_duration(str(member.id), end_time)

            print(f"{member.name} left {channel.name}. Start: {start_time}, End: {end_time}")
        
        # 退出後、チャンネル内に1人だけ残った場合は記録を停止
        if len(channel.members) == 1:
            print(f"Only one member left in {channel.name}. Stopping record for {channel.name}.")
            for member in channel.members:
                if member.id in vc_tracking:
                    # 1人だけ残った場合、記録の終了処理を追加する
                    #end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    # 日本時間を追加し、表示形式をデフォルトに変更 20250503 R.TSURUTA
                    end_time = datetime.now(jst).strftime("%Y/%m/%d %H:%M:%S")

                    update_vc_end_time_and_duration(str(member.id), end_time)
                    vc_tracking.pop(member.id)
                    print(f"Stopped recording for {member.name}.")
                vc_active[member.id] = False  # Aさんの活動状態を停止

        # 2人以上になった場合、再度記録を開始
        elif len(channel.members) >= 2:
            print(f"VC has more than 1 member. Starting record again.")
            # 日本時間を追加し、表示形式をデフォルトに変更 20250503 R.TSURUTA
            #now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            now = datetime.now(jst)
            
            for m in channel.members:
                if m.id not in vc_tracking and vc_active.get(m.id, True) is False:
                    # 新しい行を追加してそのユーザーの記録を開始
                    vc_tracking[m.id] = now
                    sheet.append_row([
                        str(m.id),
                        m.name,
                        now,
                        "VC",
                        now,
                        ""
                    ])
                    print(f"{m.name}'s record started again after others joined.")
                    vc_active[m.id] = True  # Cさんが来たタイミングで活動状態を再開

# Botの実行
bot.run(discord_token)
