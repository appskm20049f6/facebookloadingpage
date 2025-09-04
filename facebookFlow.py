
import os
from dotenv import load_dotenv
import ttkbootstrap as tb
from ttkbootstrap.dialogs import Messagebox
import requests
import random
import threading
from datetime import datetime

def show_error(title, msg):
    Messagebox.show_error(msg, title=title)

# 你的粉絲專頁權杖 (請定期更新，建議用.env檔隱藏)
load_dotenv()
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")

def fetch_posts(access_token, limit=100):
    """
    從 Facebook Graph API 取得粉絲專頁的貼文列表。
    """
    url = f"https://graph.facebook.com/v23.0/me/posts?fields=id,message,created_time&access_token={access_token}&limit={limit}"
    posts = []
    while url and len(posts) < limit:
        res = requests.get(url).json()
        if 'data' in res:
            posts.extend(res['data'])
        if 'paging' in res and 'next' in res['paging'] and len(posts) < limit:
            url = res['paging']['next']
        else:
            url = None
    return posts[:limit]

def get_all_comments(post_id, access_token):
    """
    抓取指定貼文的所有留言，並返回包含留言者姓名、網址、留言內容和時間的元組列表。
    """
    url = f"https://graph.facebook.com/v23.0/{post_id}/comments?fields=id,from,message,created_time&access_token={access_token}"
    all_comments = []
    while url:
        response = requests.get(url).json()
        if 'data' in response:
            for comment in response['data']:
                from_user = comment.get('from', {})
                commenter_name = from_user.get('name', '匿名使用者')
                
                # 根據您的要求，使用 comment 的完整 ID 來建立網址
                comment_full_id = comment.get('id', '')
                commenter_id = comment_full_id.split('_')[-1]
                profile_url = f"https://www.facebook.com/{commenter_id}" if commenter_id else ''

                comment_message = comment.get('message', '')
                comment_time = comment.get('created_time', '')
                comment_info = (commenter_name, profile_url, comment_message, comment_time)
                all_comments.append(comment_info)
            if 'paging' in response and 'next' in response['paging']:
                url = response['paging']['next']
            else:
                url = None
        else:
            url = None
    return all_comments

def start_draw():
    """
    點擊按鈕後執行，啟動抽獎流程。
    """
    post_idx = post_select.current()
    if post_idx < 0:
        show_error("錯誤", "請先選擇一則貼文！")
        return
    post_id = post_ids[post_idx]
    
    comment_keyword = entry_comment_keyword.get().strip()
    
    status_label.config(text="狀態：正在抓取留言，請稍候...")
    
    draw_thread = threading.Thread(target=draw_worker, args=(post_id, comment_keyword))
    draw_thread.start()

def draw_worker(post_id, comment_keyword):
    """
    在背景執行緒中執行留言抓取和過濾。
    """
    from datetime import datetime
    all_comments = get_all_comments(post_id, ACCESS_TOKEN)
    
    # 關鍵字過濾
    if comment_keyword:
        all_comments = [c for c in all_comments if comment_keyword in c[2]]
    
    # 圖形化時間區間過濾
    time_start = f"{date_start.get_date().strftime('%Y-%m-%d')} {hour_start.get()}:{minute_start.get()}"
    time_end = f"{date_end.get_date().strftime('%Y-%m-%d')} {hour_end.get()}:{minute_end.get()}"
    
    def parse_time(s):
        try:
            return datetime.strptime(s, "%Y-%m-%d %H:%M")
        except:
            return None
    
    dt_start = parse_time(time_start)
    dt_end = parse_time(time_end)

    if dt_start:
        all_comments = [c for c in all_comments if c[3] and datetime.strptime(c[3][:16], "%Y-%m-%dT%H:%M") >= dt_start]
    if dt_end:
        all_comments = [c for c in all_comments if c[3] and datetime.strptime(c[3][:16], "%Y-%m-%dT%H:%M") <= dt_end]

    root.after(0, lambda: perform_draw_gui(all_comments))

def perform_draw_gui(all_comments):
    """
    在抓取完留言後，執行抽獎並更新介面。
    """
    if not all_comments:
        show_error("錯誤", "無法取得留言或無符合關鍵字留言，請檢查貼文ID、權杖或關鍵字。")
        status_label.config(text="狀態：待命中")
        return
    
    total_comments = len(all_comments)
    status_label.config(text=f"狀態：成功抓取 {total_comments} 則留言。")
    
    try:
        num_winners = int(entry_winners.get())
    except ValueError:
        show_error("錯誤", "請輸入一個有效的數字。")
        status_label.config(text="狀態：待命中")
        return
    
    if num_winners <= 0 or num_winners > total_comments:
        show_error("錯誤", f"請輸入一個介於 1 到 {total_comments} 之間的數字。")
        status_label.config(text="狀態：待命中")
        return
    
    winners = random.sample(all_comments, num_winners)
    
    results_text.delete("1.0", "end")
    
    # 輸出純 tab 分隔資料，方便 Excel 複製
    header = ["中獎者姓名", "網址", "留言內容", "留言時間"]
    results_text.insert("end", "\t".join(header) + "\n")
    
    for i, (name, profile_url, message, comment_time) in enumerate(winners, 1):
        row = f"{name}\t{profile_url}\t{message.replace(chr(9),' ').replace(chr(10),' ')}\t{comment_time}\n"
        results_text.insert("end", row)
    
    status_label.config(text="狀態：抽獎完成！")

def reload_posts():
    keyword = entry_post_keyword.get().strip()
    status_label.config(text="狀態：正在載入貼文...")
    root.update()
    posts = fetch_posts(ACCESS_TOKEN, limit=100)
    filtered = [p for p in posts if keyword.lower() in (p.get('message', '') or '').lower()]
    global post_ids, post_titles
    post_ids = [p['id'] for p in filtered]
    post_titles = [f"{p.get('created_time','')[:10]} | {p.get('message','').replace('\\n',' ')[:40]}..." for p in filtered]
    post_select['values'] = post_titles
    if post_titles:
        post_select.current(0)
        status_label.config(text=f"狀態：共 {len(post_titles)} 筆貼文")
    else:
        status_label.config(text="狀態：查無符合關鍵字的貼文")

# --- GUI ---
root = tb.Window(themename="flatly")
root.title("Facebook 留言抽獎工具（可選貼文）")
root.geometry("800x700")
root.resizable(False, False)

main_frame = tb.Frame(root, padding=20)
main_frame.pack(fill="both", expand=True)

# 關鍵字搜尋欄
search_frame = tb.Frame(main_frame)
search_frame.pack(anchor="w", pady=(0, 5))
tb.Label(search_frame, text="貼文關鍵字：").pack(side="left")
entry_post_keyword = tb.Entry(search_frame, width=20)
entry_post_keyword.pack(side="left", padx=(0, 5))
reload_posts_btn = tb.Button(search_frame, text="搜尋/重載", command=reload_posts, bootstyle="info")
reload_posts_btn.pack(side="left")

# 預設先載入全部
posts = fetch_posts(ACCESS_TOKEN, limit=100)
post_ids = [p['id'] for p in posts]
post_titles = [f"{p.get('created_time','')[:10]} | {p.get('message','').replace('\\n',' ')[:40]}..." for p in posts]

tb.Label(main_frame, text="選擇要抽獎的貼文：").pack(anchor="w", pady=(0, 5))
post_select = tb.Combobox(main_frame, values=post_titles, state="readonly", width=80)
post_select.pack(anchor="w", pady=(0, 10))
if post_titles:
    post_select.current(0)

tb.Label(main_frame, text="留言關鍵字（僅抽留言內容含此字者）：").pack(anchor="w", pady=(0, 5))
entry_comment_keyword = tb.Entry(main_frame, width=20)
entry_comment_keyword.pack(anchor="w", pady=(0, 5))

# 圖形化留言時間區間（ttkbootstrap）
time_frame = tb.Frame(main_frame)
time_frame.pack(anchor="w", pady=(0, 10))

# 開始時間區塊
start_frame = tb.Frame(time_frame)
start_frame.pack(side="left", padx=(0, 40))
tb.Label(start_frame, text="留言開始時間：").pack(side="left")
date_start = tb.DateEntry(start_frame, width=12, bootstyle="info", dateformat="%Y-%m-%d")
date_start.pack(side="left")
hour_start = tb.Combobox(start_frame, width=2, values=[f"{i:02d}" for i in range(24)], state="readonly")
hour_start.set("00")
hour_start.pack(side="left")
tb.Label(start_frame, text=":").pack(side="left")
minute_start = tb.Combobox(start_frame, width=2, values=[f"{i:02d}" for i in range(60)], state="readonly")
minute_start.set("00")
minute_start.pack(side="left")

# 結束時間區塊
end_frame = tb.Frame(time_frame)
end_frame.pack(side="left")
tb.Label(end_frame, text="留言結束時間：").pack(side="left")
date_end = tb.DateEntry(end_frame, width=12, bootstyle="info", dateformat="%Y-%m-%d")
date_end.pack(side="left")
hour_end = tb.Combobox(end_frame, width=2, values=[f"{i:02d}" for i in range(24)], state="readonly")
hour_end.set("23")
hour_end.pack(side="left")
tb.Label(end_frame, text=":").pack(side="left")
minute_end = tb.Combobox(end_frame, width=2, values=[f"{i:02d}" for i in range(60)], state="readonly")
minute_end.set("59")
minute_end.pack(side="left")

# 中獎人數
tb.Label(main_frame, text="抽出人數：").pack(anchor="w", pady=(0, 5))
entry_winners = tb.Entry(main_frame, width=10)
entry_winners.pack(anchor="w", pady=(0, 10))

btn = tb.Button(main_frame, text="開始抽獎", command=start_draw, bootstyle="success")
btn.pack(pady=(10, 20))

status_label = tb.Label(main_frame, text="狀態：待命中", bootstyle="info")
status_label.pack(anchor="w", pady=(0, 10))

results_label = tb.Label(main_frame, text="抽獎結果：", bootstyle="secondary")
results_label.pack(anchor="w")

results_text = tb.Text(main_frame, wrap="word", height=15)
results_text.pack(fill="both", expand=True)

root.mainloop()
