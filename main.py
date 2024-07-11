from selenium import webdriver
from chromedriver_py import binary_path
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import re
from bs4 import BeautifulSoup
import requests
import datetime
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import os
import json

# ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

# Variables - GitHub
# line_notify_id = os.environ['LINE_NOTIFY_ID']
# sheet_key = os.environ['GOOGLE_SHEETS_KEY']
# gs_credentials = os.environ['GS_CREDENTIALS']

# Variables - Google Colab
line_notify_id = LINE_NOTIFY_ID
sheet_key = GOOGLE_SHEETS_KEY
gs_credentials = GS_CREDENTIALS

# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

# LINE Notify ID
LINE_Notify_IDs = list(line_notify_id.split())

# 定義查找nid代碼函數
def find_nid(title, text):
    title_line_numbers = []
    for i, line in enumerate(text.split('\n')):
        if title in line:
            title_line_numbers.append(i)

    if not title_line_numbers:
        print(f'Cannot find "{title}" in the text.')
        return None

    title_line_number = title_line_numbers[0]
    title_line = text.split('\n')[title_line_number]

    nid_start_index = title_line.index('nid="') + 5
    nid_end_index = title_line.index('"', nid_start_index)
    nid = title_line[nid_start_index:nid_end_index]

    return nid


def download_file(url, folder_path, idx):
    # 解析URL獲取文件名
    parsed_url = urlparse(url)
    file_name = os.path.join(folder_path, f'{idx:03} - {unquote(os.path.basename(parsed_url.path))}')
    
    # 發送請求下載文件
    response = requests.get(url)
    with open(file_name, 'wb') as f:
        f.write(response.content)
    
    # print(f"文件已下載到 {file_name}")

# 取得網頁內容
def get_content(url, send):

  # chromedriver 設定
  options = webdriver.ChromeOptions()
  options.add_argument('--headless')
  options.add_argument('--no-sandbox')
  options.add_argument('--disable-dev-shm-usage')

  service = Service(binary_path)
  driver = webdriver.Chrome(service=service, options=options)
  driver.get(url)

  # 等待網頁載入完成
  driver.implicitly_wait(10)

  # 發送GET請求獲取網頁內容
  response = requests.get(url)

  # 打印出網頁的HTML內容
  page_html = driver.page_source
  # print(page_html)
  # driver.quit()

  # 使用BeautifulSoup解析HTML
  soup = BeautifulSoup(page_html, 'html.parser')

  if send:
    # 提取附件連結和檔案名稱
    attachments = []
    attach_list_divs = soup.find_all('div', id='attach-list')

    for div in attach_list_divs:
      link_tag = div.find('a')
      if link_tag and 'href' in link_tag.attrs:
        href = link_tag['href']
        text = link_tag.text.strip()
        attachments.append({'name': text, 'url': href})

    if not os.path.exists(input_path):
      os.makedirs(input_path)

    empty = True
    idx = 1
    # 打印附件連結和檔案名稱
    for attachment in attachments:
      empty = False
      attachment_name = attachment['name']
      attachment_url = attachment['url']
      print(f"Name: {attachment_name}, URL: {attachment_url}")
    
      download_file(attachment_url, input_path, f"{idx:03}")
      idx += 1
  else:
    empty = True

  # 解析HTML內容
  soup = BeautifulSoup(response.content, 'html.parser')
  # print(soup)

  # 格式化HTML文件
  formatted_html = soup.prettify()
  # print(formatted_html)

  # 找到所有的 <p> 標籤
  p_tags = soup.find_all('p')

  # 整理文字內容
  text_list = []
  for p in p_tags:
    text = p.text.strip()
    text_list.append(text)
  text = ' '.join(text_list)
  text = ' '.join(text.split())  # 利用 split() 和 join() 將多個空白轉成單一空白
  # text = text.replace(' ', '\n')  # 將空白轉換成換行符號
  text = text.replace(' ', '')  # 刪除空白
  return text, empty

# 假設 input_path 是資料夾路徑，包含要刪除的檔案
def delete_files_in_folder(input_path):
    # 檢查資料夾是否存在
    if not os.path.exists(input_path):
        # print(f"資料夾 {input_path} 不存在。")
        return
    
    # 刪除資料夾內的所有檔案
    for filename in os.listdir(input_path):
        file_path = os.path.join(input_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                # print(f"刪除 {file_path} 成功。")
        except Exception as e:
            # print(f"刪除 {file_path} 時發生錯誤：{str(e)}")
            pass

def convert_files_to_images(input_path, output_path, max_pages=4):
    # 檢查輸出資料夾是否存在，若不存在則建立
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    
    # 儲存所有圖片的路徑
    all_image_files = []

    # 處理資料夾中的每個檔案
    for filename in os.listdir(input_path):
        file_path = os.path.join(input_path, filename)
        
        # 如果是PDF檔案，則轉換為PNG圖片
        if filename.lower().endswith('.pdf'):
            image_files = pdf_to_png(file_path, output_path, max_pages)
            all_image_files.extend(image_files)
            if image_files:
                pass
                # print(f"成功將 {filename} 轉換並合併為一張圖片。")
        
        # 如果是圖片檔案（JPG或PNG），直接複製到輸出資料夾
        elif filename.lower().endswith(('.jpg', '.jpeg', '.png')):
            img = Image.open(file_path)
            img.save(os.path.join(output_path, f"{filename.split('.')[0]}.png"), 'PNG')
            all_image_files.append(os.path.join(output_path, f"{filename.split('.')[0]}.png"))
            # print(f"成功複製 {filename} 到輸出資料夾。")

    # 如果有圖片檔案，則合併成一張圖片
    if all_image_files:
        global combined_image_path
        combined_image_path = os.path.join(output_path, "combined_image.jpg")
        combine_images(all_image_files, combined_image_path)
        # print(f"所有圖片已成功合併為 combined_image.jpg。")

        # 刪除暫存圖片
        for file in all_image_files:
            os.remove(file)
            # print(f"刪除 {file}。")

def pdf_to_png(pdf_path, output_path, max_pages=4):
    # 將PDF的前四頁（或更少）轉換為PNG圖片
    images = convert_from_path(pdf_path, first_page=1, last_page=max_pages)
    
    # 儲存每一頁為單獨的PNG文件
    image_files = []
    for i, image in enumerate(images):
        output_path = f"{output_path}/{os.path.basename(pdf_path).split('.')[0]}_page_{i + 1}.png"
        image.save(output_path, 'PNG')
        image_files.append(output_path)
    
    return image_files

def image_to_a4(image):

    A4_WIDTH = 2480//2
    A4_HEIGHT = 3508//2

    # 將圖片縮放至A4大小
    target_width = A4_WIDTH
    target_height = A4_HEIGHT
    width, height = image.size
    
    # 計算縮放比例
    ratio_width = target_width / width
    ratio_height = target_height / height
    ratio = min(ratio_width, ratio_height)
    
    # 調整圖片大小並居中截取
    new_width = int(width * ratio)
    new_height = int(height * ratio)
    resized_image = image.resize((new_width, new_height), Image.LANCZOS)
    x_offset = (target_width - new_width) // 2
    y_offset = (target_height - new_height) // 2
    
    # 創建新的白色背景圖片
    background = Image.new('RGB', (target_width, target_height), (255, 255, 255))
    background.paste(resized_image, (x_offset, y_offset))
    
    return background

def combine_images(image_files, output_path):
    num_images = len(image_files)
    
    if num_images == 0:
        return
    
    # 開啟所有圖片
    images = [image_to_a4(Image.open(img)) for img in image_files]
    
    # # 初始化合併後的圖片變量
    # combined_image = None
    
    if num_images == 1:
        # 1x1 layout
        combined_image = images[0]
    elif num_images == 2:
        # 1x2 layout
        widths, heights = zip(*(img.size for img in images))
        total_width = sum(widths)
        max_height = max(heights)
        combined_image = Image.new('RGB', (total_width, max_height))
        x_offset = 0
        for img in images:
            combined_image.paste(img, (x_offset, 0))
            x_offset += img.width
    elif num_images == 3:
        # 1x3 layout
        widths, heights = zip(*(img.size for img in images))
        total_width = sum(widths)
        max_height = max(heights)
        combined_image = Image.new('RGB', (total_width, max_height))
        x_offset = 0
        for img in images:
            combined_image.paste(img, (x_offset, 0))
            x_offset += img.width
    elif num_images == 4:
        # 2x2 layout
        widths, heights = zip(*(img.size for img in images))
        total_width = max(widths) * 2
        total_height = max(heights) * 2
        combined_image = Image.new('RGB', (total_width, total_height))
        positions = [(0, 0), (widths[0], 0), (0, heights[0]), (widths[0], heights[0])]
        for pos, img in zip(positions, images):
            combined_image.paste(img, pos)

    # 儲存合併後的圖片
    if combined_image is not None:
      combined_image.save(output_path)

# LINE Notify
def LINE_Notify(category, date, title, unit, link, content, empty):
  
  text_limit = 1000-3

  send_info_1 = f'【{category}】{title}\n⦾公告日期：{date}\n⦾發佈單位：{unit}'
  send_info_2 = f'⦾內容：' if content != '' else ''
  send_info_3 = f'⦾更多資訊：{link}'

  text_len = len(send_info_1) + len(send_info_2) + len(send_info_3)
  if content != '':
    if text_len + len(content) > text_limit:
      content = f'{content[:(text_limit - text_len)]}⋯'
    params_message = f'{send_info_1}\n{send_info_2}{content}\n{send_info_3}'
  else:
    params_message = f'{send_info_1}\n{send_info_3}'

  for LINE_Notify_ID in LINE_Notify_IDs:
    
    headers = {
            'Authorization': 'Bearer ' + LINE_Notify_ID,
            # 'Content-Type': 'application/x-www-form-urlencoded'
        }
    params = {'message': params_message}

    if empty:
      r = requests.post('https://notify-api.line.me/api/notify',
                              headers=headers, params=params)
    else:
      # print(f'{output_path}/combined_image.jpg')
      image = open(f'{output_path}/combined_image.jpg', 'rb')
      # image = open('/content/digital_camera_photo-1080x675.jpg', 'rb')
      files = { 'imageFile': image }
      r = requests.Session().post('https://notify-api.line.me/api/notify',
                              headers=headers, data=params, files=files)
      print('photo', end=' ')

    # print(r.text)  #200
    print(r.status_code)  #200

# Google Sheets 紀錄
scope = ['https://www.googleapis.com/auth/spreadsheets']
info = json.loads(gs_credentials)

creds = Credentials.from_service_account_info(info, scopes=scope)
gs = gspread.authorize(creds)

def google_sheets_refresh():

  global sheet, worksheet, rows_sheets, df

  # 使用表格的key打開表格
  sheet = gs.open_by_key(sheet_key)
  worksheet = sheet.get_worksheet(0)

  # 讀取所有行
  rows_sheets = worksheet.get_all_values()
  # 使用pandas創建數據框
  df = pd.DataFrame(rows_sheets)

def main():

  # 開啟網頁
  urls = {
      '最新消息':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_0175a41dca498eab35e73c7c40fd1c141d1f3a58&maximize=1&allbtn=0' # 最新消息
      # ,'榮譽榜':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_0f31e8a5a7bc3c4ef6609345e33cd3ae6b3e97cc&maximize=1&allbtn=0' # 榮譽榜
      # ,'校務通報':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b1568f22c498d41d46e53b69325a31e78d51c87c&maximize=1&allbtn=0' # 校務通報
      # ,'研習':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_e04b673ba655cf5fbdbab4e815d941418b3ec90c&maximize=1&allbtn=0'# 研習
      # ,'師生活動與競賽':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_548130b3d109474559de5f5f564d0a729c3c7b3d&maximize=1&allbtn=0' # 師生活動與競賽
      # ,'防疫專區':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_36ca124c0de704b62aead8039c9bbe125095f5aa&maximize=1&allbtn=0' # 防疫專區
      # ,'招標公告':'https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_27c842d8808109d6838dde8f0c1222d6177d71b8&maximize=1&allbtn=0' #招標公告
  }

  # 刷新Google Sheets表格
  google_sheets_refresh()

  # 取得Google Sheets nids列表
  _nids = df[5].tolist()
  nids = []
  for n in _nids:
    try:
      nids.append(str(int(n)))
    except:
      continue

  for category in urls:

      url = urls[category]

      # chromedriver 設定
      options = webdriver.ChromeOptions()
      options.add_argument('--headless')
      options.add_argument('--no-sandbox')
      options.add_argument('--disable-dev-shm-usage')

      service = Service(binary_path)
      driver = webdriver.Chrome(service=service, options=options)
      driver.get(url)

      # 等待網頁載入完成
      driver.implicitly_wait(10)

      # 找到表格元素
      res = requests.get(url)
      soup = BeautifulSoup(res.text, 'html.parser')
      table_div = driver.find_element(By.ID, 'div_table_content')
      table = table_div.find_element(By.ID, 'ntb')
      html = table.get_attribute('outerHTML')

      # 解析HTML文件
      soup = BeautifulSoup(html, 'html.parser')

      # 格式化HTML文件
      formatted_html = soup.prettify()
      # print(formatted_html)

      # 找到表格中的所有資料列
      rows = table.find_elements(By.TAG_NAME, 'tr')

      # 打印每一行的 HTML 內容
      # for row in rows:
      #     row_html = row.get_attribute('outerHTML')
      #     print(row_html)

      # 定義需要查找的最新幾筆資料（最多9筆）
      numbers_of_new_data = 9
      # print(len(rows)-1)
      numbers_of_new_data = min(numbers_of_new_data, len(rows)-1)

      # 印出最新幾筆資料的標題、單位和連結
      for i in range(numbers_of_new_data - 1, -1, -1):
          row = rows[numbers_of_new_data - i]
          # row_html = row.get_attribute('outerHTML')
          # print(row_html)
          cells = row.find_elements(By.TAG_NAME, 'td')
          date = cells[1].text
          title = cells[3].text
          unit = cells[2].text

          # 使用 BeautifulSoup 解析 HTML
          soup = BeautifulSoup(row.get_attribute('outerHTML'), 'html.parser')

          # 找到 nid 的值
          nid = soup.find('tr')['nid']

          # 檢查nid是否已經存在於表格中
          send = not(str(int(nid)) in nids)

          link_publish = f'http://www.kjsh.ntpc.edu.tw/ischool/public/news_view/show.php?nid={nid}'
          link = f'lihi.cc/depwP/{nid}'
          content, empty = get_content(link_publish, send)
          print(f'date:{date}\tcategory:{category}\ttitle:{title}\tunit:{unit}\tnid:{nid}\tlink:{link}\tcontent:{content}')

          if not empty:

            # 將資料夾中的檔案轉換成圖片並合併
            convert_files_to_images(input_path, output_path)

            # print(f"所有檔案轉換並合併完成，並刪除了暫存圖片。")

            # 刪除input檔案
            delete_files_in_folder(input_path)


          # 獲取當前日期
          today = datetime.date.today()

          # 將日期格式化為2023/02/11的形式
          formatted_date = today.strftime("%Y/%m/%d")


          if send:

            # 檢查標題是否已經存在於表格中
            titles = df[3].tolist()
            if title in titles:
              continue

            # 獲取新行
            now = datetime.datetime.now() + datetime.timedelta(hours=8)
            new_row = [now.strftime("%Y-%m-%d %H:%M:%S"), category, date, title, unit, nid, link, content]

            # 將新行添加到工作表中
            worksheet.append_row(new_row)

            # 獲取新行的索引
            new_row_index = len(rows) + 1

            # 更新單元格
            cell_list = worksheet.range('A{}:H{}'.format(new_row_index, new_row_index))
            for cell, value in zip(cell_list, new_row):
                cell.value = value
            worksheet.update_cells(cell_list)

            # 更新nids列表
            nids.append(int(nid))

            # 傳送至LINE Notify
            print(f'send: {nid}', end=' ')
            LINE_Notify(category, date, title, unit, link, content, empty)

          if not empty:
            # 刪除合併後的圖片
            # if not os.path.exists(output_path):
            os.remove(combined_image_path)
            # print(f"刪除 {combined_image_path}。")

          # 刪除nid
          del nid

      # 關閉網頁
      driver.quit()

input_path = './input'
output_path = './output'

if __name__ == "__main__":

  try_times_limit = 2
  for _ in range(try_times_limit):
    try:
      main()
      break
    except:
      next
