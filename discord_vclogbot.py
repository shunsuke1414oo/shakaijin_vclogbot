import discord
from discord.ext import commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json
<<<<<<< HEAD
import pytz

# 設定ファイルの読み込み
=======
import pytz #20250503 R.TSURUTA(日本時間実装のため)
# 設定ファイルを読み込む
>>>>>>> b769a14dcdaf1e72bf0d34d40b17c97df9b8c2f6
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# 設定値の取得
discord_token = config["discord_token"]
google_creds_file = config["google_creds_file"]
spreadsheet_name = config["spreadsheet_name"]

# Discordのインテント設定
intents = discord.Intents.default()
intents.voice_states = True
intents.members = True
intents.message_content = True 

# JSTのタイムゾーン
jst = pytz.timezone('Asia/Tokyo')

<<<<<<< HEAD
# Google Sheets 認証と接続
=======
# 日本標準時 (JST) のタイムゾーンを指定　20250503　R.TSURUTA
jst = pytz.timezone('Asia/Tokyo')

# Google Sheets認証
>>>>>>> b769a14dcdaf1e72bf0d34d40b17c97df9b8c2f6
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(google_creds_file, scope)
client = gspread.authorize(creds)
sheet = client.open(spreadsheet_name).sheet1

# VCの状態管理用
vc_tracking = {} # ユーザーID: VC参加時刻 (datetimeオブジェクト)
vc_active = {}   # ユーザーID: 現在VCで活動中か (True/False)

# Botクラスの定義とログ出力
class MyBot(commands.Bot):
    def dispatch(self, event_name, *args, **kwargs):
        print(f"[Event] {event_name} triggered with args={args} kwargs={kwargs}")
        super().dispatch(event_name, *args, **kwargs)

bot = MyBot(command_prefix="!", intents=intents)

# VC開始の記録処理
def record_vc_start(member, timestamp):
    formatted_time = timestamp.strftime("%Y/%m/%d %H:%M:%S")
    sheet.append_row([
        str(member.id),
        member.name,
        formatted_time, # C列: イベント発生時刻 (実際にはVC参加を検知した時刻)
        "VC",           # D列: イベントタイプ
        formatted_time, # E列: VC開始時刻
        ""              # F列: VC終了時刻 (最初は空)
        # G列（滞在時間）は update_vc_end_time_and_duration で数式が入力される
    ])
    vc_tracking[str(member.id)] = timestamp # ユーザーIDは文字列で統一
    vc_active[str(member.id)] = True
    print(f"{member.name} joined VC at {formatted_time}")

# VC終了の記録処理
def update_vc_end_time_and_duration(member_id: str, end_time_obj: datetime):
    formatted_time = end_time_obj.strftime("%Y/%m/%d %H:%M:%S")
    rows = sheet.get_all_values()
    target_row_f = None # 終了時刻(F列)を更新する行番号
    target_row_g = None # 滞在時間(G列)を更新する行番号 (F列と同じ行のはず)

    # スプレッドシートを下から検索して、該当ユーザーの最新の未終了VC記録を探す
    for i in range(len(rows) - 1, 0, -1): # ヘッダー行を避けるため0まで (1行目からデータ想定)
        row = rows[i]
<<<<<<< HEAD
        # ユーザーIDが一致し、イベントタイプが"VC"である行を探す
        if row[0] == member_id and row[3] == "VC":
            # F列 (終了時刻) が空欄の行を見つけたら、その行を更新対象とする
            if len(row) > 5 and row[5] == "" and target_row_f is None: # row[5]はF列 (IndexError対策でlenも見る)
                target_row_f = i + 1 # gspreadの行番号は1から始まるため +1
            # G列 (滞在時間) が空欄の行を見つけたら、その行を更新対象とする
            # (実際にはF列と同じ行になるはずだが、念のため分けている)
            if len(row) > 6 and row[6] == "" and target_row_g is None: # row[6]はG列 (IndexError対策でlenも見る)
                target_row_g = i + 1
            # 両方の更新対象行が見つかればループを抜ける (F列かG列のどちらかの更新対象が見つかれば良い、という考え方もできる)
            # 今回はF列が空の最新行にFとGをセットするので、target_row_fが見つかればtarget_row_gも同じ行になる
            if target_row_f: # target_row_f が見つかれば、target_row_g も同じ行とみなして良い
                if target_row_g is None : target_row_g = target_row_f # もしG列の判定がうまく行かなくてもF列に合わせる
                break
            # もし厳密にF列もG列も空の行を探すなら if target_row_f and target_row_g: break
=======
        
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
>>>>>>> b769a14dcdaf1e72bf0d34d40b17c97df9b8c2f6
    
    updates = []
    if target_row_f:
        updates.append({
            'range': f'F{target_row_f}',
            'values': [[formatted_time]]
        })
    else:
        print(f"No matching row found for user {member_id} with empty end time to update F column.")

    if target_row_g: # F列を更新した行と同じ行のG列を更新する
        duration_formula = f'=IF(AND(ISBLANK(E{target_row_g}),ISBLANK(F{target_row_g})),"",IF(OR(ISBLANK(E{target_row_g}),ISBLANK(F{target_row_g})),"",INT((F{target_row_g}-E{target_row_g})*86400)))'
        updates.append({
            'range': f'G{target_row_g}',
            'values': [[duration_formula]]
        })
    else:
        # target_row_f が見つからなければ target_row_g も見つからないはず
        print(f"No matching row found for user {member_id} with empty duration to update G column.")


    if updates:
        # 複数のセルを一括で更新
        body = {'valueInputOption': 'USER_ENTERED', 'data': updates}
        sheet.spreadsheet.values_batch_update(body=body)
        print(f"Updated VC end and duration for user {member_id} at row {target_row_f if target_row_f else 'N/A'}.")
    else:
        print(f"No updates made for user {member_id}.")

# VCステート変更イベント
@bot.event
async def on_voice_state_update(member, before, after):
<<<<<<< HEAD
    now = datetime.now(jst)
    user_id_str = str(member.id) # ユーザーIDを文字列で保持

    # デバッグ用ログ (必要に応じてコメント解除)
    # print(f"[VSU Debug] User: {member.name} ({user_id_str}), Before: {before.channel}, After: {after.channel}")
    # print(f"[VSU Debug] vc_tracking before: {vc_tracking.get(user_id_str)}, vc_active before: {vc_active.get(user_id_str)}")
=======
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
>>>>>>> b769a14dcdaf1e72bf0d34d40b17c97df9b8c2f6

    # チャンネル移動、参加、退出の判定
    is_join = before.channel is None and after.channel is not None
    is_leave = before.channel is not None and after.channel is None
    is_move = before.channel is not None and after.channel is not None and before.channel != after.channel

    # --- 1. 退出関連処理 ---
    # 対象: 純粋な退出 (is_leave) または 移動の「元」チャンネルからの退出 (is_move)
    if is_leave or is_move:
        # イベントを発生させたメンバーが記録対象だった場合、そのセッションを終了する
        if user_id_str in vc_tracking: 
            print(f"[VC Processing] Leave/Move-out: {member.name} from {before.channel.name}")
            update_vc_end_time_and_duration(user_id_str, now) 

<<<<<<< HEAD
            # 追跡情報をクリア/更新
            vc_tracking.pop(user_id_str, None) 
            vc_active[user_id_str] = False     
            print(f"{member.name} left VC or moved from {before.channel.name}.")
        # else:
            # print(f"[VC Info] {member.name} left/moved from {before.channel.name}, but was not actively tracked in vc_tracking.")

        # チャンネルに残ったメンバーの処理 (1人だけになった場合の処理)
        # この処理は、`member` が `before.channel` を抜けた「後」のメンバーリストで評価される
        if len(before.channel.members) == 1:
            remaining_member = before.channel.members[0]
            remaining_id_str = str(remaining_member.id)
            if remaining_id_str in vc_tracking: # 残ったメンバーが記録中だったら
                print(f"[VC Processing] Last member logic: {remaining_member.name} in {before.channel.name}")
                update_vc_end_time_and_duration(remaining_id_str, now) 
                vc_tracking.pop(remaining_id_str, None)
                vc_active[remaining_id_str] = False
                print(f"{remaining_member.name}'s recording stopped (only one left in {before.channel.name}).")

    # --- 2. 入室関連処理 ---
    # 対象: 純粋な参加 (is_join) または 移動の「先」チャンネルへの参加 (is_move)
    if is_join or is_move:
        target_channel = after.channel # 参加/移動先のチャンネル
        member_count = len(target_channel.members) 
=======
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
>>>>>>> b769a14dcdaf1e72bf0d34d40b17c97df9b8c2f6

        # チャンネルに2人以上いる場合のみ記録処理を行う
        if member_count >= 2:
            # A. イベントを発生させたメンバー自身の記録開始/再開
            # 移動の場合、上記「退出関連処理」で vc_tracking からIDがpopされ、vc_activeがFalseになっているはず。
            if user_id_str not in vc_tracking or not vc_active.get(user_id_str, False):
                print(f"[VC Processing] Join/Move-in: Recording start for {member.name} in {target_channel.name}")
                record_vc_start(member, now) 
            # else:
                # print(f"[VC Info] {member.name} is already tracked and active in {target_channel.name}, skipping record_vc_start.")

<<<<<<< HEAD
            # B. チャンネル内の他のメンバーで、まだ記録されていないか非アクティブな人がいれば記録開始/再開
            for m_in_channel in target_channel.members:
                m_id_str = str(m_in_channel.id)
                if m_id_str != user_id_str: # イベントを発生させた本人以外のメンバーをチェック
                    if m_id_str not in vc_tracking or not vc_active.get(m_id_str, False):
                        print(f"[VC Processing] Join/Move-in: Recording start for other member {m_in_channel.name} in {target_channel.name} (channel count: {member_count})")
                        record_vc_start(m_in_channel, now)
        # else: # チャンネルのメンバーが1人だけの場合
            # 純粋な参加で1人だけVCに入った場合は、元々記録対象外。
            # 移動の結果1人になった場合、その人の記録は移動元の処理で既に閉じられているはず。
            # print(f"[VC Info] {member.name} is in {target_channel.name} alone or with <2 members. No new recording started by this event.")

    # デバッグ用ログ (必要に応じてコメント解除)
    # print(f"[VSU Debug] vc_tracking after: {vc_tracking.get(user_id_str)}, vc_active after: {vc_active.get(user_id_str)}")

# Bot起動
bot.run(discord_token)
=======
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
>>>>>>> b769a14dcdaf1e72bf0d34d40b17c97df9b8c2f6
