from selenium import webdriver
from bs4 import BeautifulSoup
import collections
collections.Callable = collections.abc.Callable
import openpyxl, threading, requests, ctypes, time, os, random
import warnings
warnings.filterwarnings(action='ignore')

threads = []
No = 1
smartstorelist = []
detail = {}
test = ""
lastlist = []

ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'
headers = {
    'User-Agent': ua
}
############################################################################################################################################

def listcheck(pages):
    for line in open('./resources/config/category.txt', encoding='utf-8'):
        if line.startswith('#'):
            pass
        else:
            line = line.split(' ')
            category = line[0]
            code = line[1].replace('\n', '')
            print(f'코드: {code} 카테고리: {category}')
            th = threading.Thread(target=get_smartstore, args=(code, pages))
            th.start()
            threads.append(th)
            time.sleep(3)
    for th in threads:
        th.join()
    print("Finish All threads")



def scrolldown(driver):
    scroll_location = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(0.1)
        scroll_height = driver.execute_script("return document.body.scrollHeight")
        if scroll_location == scroll_height:
            break
        else:
            scroll_location = driver.execute_script("return document.body.scrollHeight")
############################################################################################################################################

def get_smartstore(code, pages):
    global No
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument(f'--user-agent={ua}')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    if Isproxy:
        options.add_argument('--proxy-server={0}'.format(str(random.choice(list(proxies)))))
    driver = webdriver.Chrome(options=options, executable_path='./resources/driver/chromedriver.exe')
    for i in range(pages):
        url = f'https://search.shopping.naver.com/search/category/{code}?pagingIndex={str(i+1)}&productSet=overseas'
        driver.get(url)
        driver.implicitly_wait(10)
        currenturl = driver.current_url
        if currenturl == "stopit":
            if Isproxy:
                get_smartstore(code, pages)
            else:
                auto_save()
        scrolldown(driver)
        html = driver.page_source
        bs = BeautifulSoup(html, 'html.parser')
        product_list = bs.select("li[class^='basicList_item']")
        for li in product_list:
            for goods in li.contents:
                storelink = goods.select("a[class^='basicList_mall__BC5Xu']")[0]['href']
                smartstore = goods.select("a[class^='basicList_mall__BC5Xu']")[0].text #스마트스토어이름
                if smartstore.rstrip() == "쇼핑몰별 최저가":
                    driver.get(storelink)
                    driver.implicitly_wait(10)
                    currenturl = driver.current_url
                    _html = driver.page_source
                    _bs = BeautifulSoup(_html, 'html.parser')
                    _product_list = _bs.select("table[class^='productByMall_list_seller'] > tbody")
                    for a in _product_list:
                        for _goods in a.contents:
                            _storelink = _goods.select("a[class^='productByMall_mall__SIa50']")[0]['href']
                            _smartstore = _goods.select("a[class^='productByMall_mall__SIa50']")[0].text # 스마트스토어 이름
                            response = requests.get(_storelink, headers=headers)
                            if "smartstore.naver.com" in response.url:
                                _storelink = response.url
                                #print(_storelink + " - 가격비교")
                                _storelink = _storelink.split('/product')[0]
                                threading.Thread(target=_add_smartstore_list, args=(_smartstore, _storelink)).start()
                                continue
                if smartstore.rstrip() == "":
                    smartstore = goods.select("a[class^='basicList_mall__BC5Xu'] > img")[0]['alt'] #스마트스토어 이미지 이름
                if "smartstore.naver.com" in storelink:
                    #print(storelink + " - 기존")
                    threading.Thread(target=add_smartstore_list, args=(smartstore, storelink)).start()
    driver.quit()

def add_smartstore_list(smartstore, storelink):
    global No
    if smartstore not in smartstorelist:
        smartstorelist.append(smartstore)
        storecode = storelink.split('smartstore.naver.com%2F')[1].split('&')[0]
        print("기존 : " + storecode)
        detail[smartstore] = storecode

def _add_smartstore_list(smartstore, storelink):
    global No
    if smartstore not in smartstorelist:
        smartstorelist.append(smartstore)
        storecode = storelink.split('smartstore.naver.com/')[1].split('&')[0]
        print("가격비교 : " + storecode)
        detail[smartstore] = storecode

############################################################################################################################################

def auto_save():
    total, success, fail = 0, 0, 0
    for line in open('./resources/temp.txt'):
        total += 1
    wb = openpyxl.Workbook() 
    sheet = wb.active
    sheet.append(["No", "구분", "몰이름", "썸네일링크", "상품명", "카테고리", "링크", "가격", "배송비", "원산지", "리뷰개수", "리뷰평점", "리뷰1", "리뷰2", "리뷰3", "리뷰4", "리뷰5", "리뷰6", "리뷰7", "리뷰8", "리뷰9", "리뷰10", "Q&A개수", "Q&A1", "Q&A2", "Q&A3", "Q&A4", "Q&A5", "Q&A6", "Q&A7", "Q&A8", "Q&A9", "Q&A10" ])
    for line in open('./resources/temp.txt'):
        if line != "":
            try:
                success += 1
                sheet.append([str(success), line.split('###')[1], line.split('###')[2], line.split('###')[3], line.split('###')[4], line.split('###')[5], line.split('###')[6], line.split('###')[7], line.split('###')[9], line.split('###')[10], line.split('###')[11], line.split('###')[12], line.split('###')[13], line.split('###')[14], line.split('###')[15], line.split('###')[16], line.split('###')[17], line.split('###')[18], line.split('###')[19], line.split('###')[20], line.split('###')[21], line.split('###')[22], line.split('###')[23], line.split('###')[24], line.split('###')[25], line.split('###')[26], line.split('###')[27], line.split('###')[28], line.split('###')[29], line.split('###')[30], line.split('###')[31], line.split('###')[32], line.split('###')[33]])
            except:
                fail += 1
                pass
    for cell in sheet["G"]:
        if cell.value == "링크":
            pass
        else:
            cell.hyperlink = cell.value
            cell.style = "Hyperlink"
    wb.save(f"./result.xlsx")
    msg = ctypes.windll.user32.MessageBoxW(None, f'결과 => 전체[{total}] 성공[{success}] 실패[{fail}]', "크롤링 완료!", 0)
    pid = os.getpid()
    os.kill(pid, 2)

def smartstorecrawling(detail): #쓰레드 시작
    thlist = []
    for store in detail:
        th = threading.Thread(target=get_goods_in_smartstore, args=(detail, store))
        th.start()
        thlist.append(th)
        time.sleep(3)
    for th in thlist:
        th.join()

def get_goods_in_smartstore(detail, store):  #스마트스토어 안에 상품 정보 가져옴
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument(f'--user-agent={ua}')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(options=options, executable_path='./resources/driver/chromedriver.exe')
    for i in range(10):
        driver.get(f'https://smartstore.naver.com/{detail[store]}/category/ALL??st=REVIEW&page={i+1}')
        driver.implicitly_wait(10)
        html = driver.page_source
        bs = BeautifulSoup(html, 'html.parser')
        product_list = bs.select("li[class^='-qHwcFXhj0']")
        for li in product_list:
            li = str(li)
            title = li.split('<strong class="QNNliuiAk3">')[1].split('<')[0]
            price = li.split('<span class="nIAdxeTzhx">')[1].split('<')[0]
            link = 'https://smartstore.naver.com'+li.split('href="')[1].split('"')[0]
            image = li.split('class="_25CKxIKjAk" src="')[1].split('"')[0]
            threading.Thread(target=get_goods_info, args=(link, title, price, store, image)).start() #상품 상세 정보 가져오기
    driver.quit()

def get_goods_info(url, title, price, store, image):
    global No
    try:
        reviewtemp = ""
        qnatemp = ""
        r = requests.get(url).text
        if "해외직배송 상품" not in r:
            return
        storeid = r.split('"payReferenceKey":"')[1].split('"')[0] #스토어아이디
        productid = url.split('products/')[1] #제품아이디
        productno = r.split('productNo":"')[1].split('"')[0]
        origin = r.split('"원산지":"')[1].split('"')[0] #원산지
        reviewscore = r.split('사용자 총 평점</span><strong class="_2pgHN-ntx6">')[1].split('<')[0] #평균리뷰점수
        reviewc = requests.get(f'https://smartstore.naver.com/i/v1/reviews/attaches/ids-count?merchantNo={storeid}&originProductNo={productno}&reviewServiceType=SELLBLOG&sortType=REVIEW_RANKING', headers=headers).text
        review = requests.get(f'https://smartstore.naver.com/i/v1/reviews/paged-reviews?page=1&pageSize=20&merchantNo={storeid}&originProductNo={productno}&sortType=REVIEW_CREATE_DATE_DESC', headers=headers).text
        qna = requests.get(f'https://smartstore.naver.com/i/v1/comments/PRODUCTINQUIRY/{productid}?filterMyComment=false&page=1&size=10&sellerAnswerYn=true', headers=headers).text
        #상품등록일 ( 트래픽 제한으로 막힘 ) #dd = requests.get(f'https://search.shopping.naver.com/search/all?frm=NVSHOVS&origQuery=test&pagingIndex=1&pagingSize=40&productSet=overseas&query={se}&sort=rel&timestamp=&viewType=list', headers=headers).text.split('asicList_etc__2uAYO">등록일 <!-- -->')[1].split('<')[0] #등록일
        cat = r.split(',"category":"')[1].split('"')[0]
        if reviewc == "":
            return
        reviewcount = reviewc.split('{"totalCount":')[1].split(',')[0]
        qnacount = qna.split('"totalElements":')[1].split(',')[0] #qna개수
        deliv = ""
        try:
            deliv = r.split('택배배송</span><span class="bd_ChMMo"><span class="bd_3uare">')[1].split('<')[0]#배송비
        except:
            deliv = "무료배송"
        if int(reviewcount) >= 2 or int(qnacount) >= 10: #만약 리뷰와 qna가 10개 이상이면
            for i in range(10):
                try:
                    reviewdate = review.split('reviewScore')[i+1].split('createDate":"')[1].split('T')[0] #리뷰날짜
                    reviewtemp += reviewdate + "/"
                except:
                    reviewtemp += "-/"
                try:
                    qnadate = qna.split('regDate":"')[i+1].split('T')[0] #qna 날자
                    qnatemp += qnadate + "/"
                except:
                    qnatemp += "-/"
        else:
            return
        print(title)
        reviewdates = reviewtemp.split('/')
        qnadates = qnatemp.split('/')
        try:
            open('./resources/temp.txt', 'a').write(f'''{str(No)}###스마트스토어###{str(store)}###{str(image)}###{str(title)}###{str(cat)}###{str(url)}###{str(price)}###test###{str(deliv)}###{str(origin)}###{str(reviewcount)}###{str(reviewscore)}###{str(reviewdates[0])}###{str(reviewdates[1])}###{str(reviewdates[2])}###{str(reviewdates[3])}###{str(reviewdates[4])}###{str(reviewdates[5])}###{str(reviewdates[6])}###{str(reviewdates[7])}###{str(reviewdates[8])}###{str(reviewdates[9])}###{str(qnacount)}###{str(qnadates[0])}###{str(qnadates[1])}###{str(qnadates[2])}###{str(qnadates[3])}###{str(qnadates[4])}###{str(qnadates[5])}###{str(qnadates[6])}###{str(qnadates[7])}###{str(qnadates[8])}###{str(qnadates[9])}\n''')
            No += 1
        except:
            pass
    except:
        pass

if __name__ == '__main__':
    open('./resources/temp.txt', 'w')
    Isproxy = False
    ip = input("프록시를 사용하시겠습니까? (y/n) :")
    if ip == "y":
        Isproxy = True
        proxies = open("./proxy.txt", 'r').read().split('\n')
    listcheck(int(input("크롤링할 페이지 수: ")))
    print(detail)
    smartstorecrawling(detail)
    total, success, fail = 0, 1, 0
    for line in open('./resources/temp.txt'):
        total += 1
    wb = openpyxl.Workbook() 
    sheet = wb.active
    sheet.append(["No", "구분", "몰이름", "썸네일링크", "상품명", "카테고리", "링크", "가격", "배송비", "원산지", "리뷰개수", "리뷰평점", "리뷰1", "리뷰2", "리뷰3", "리뷰4", "리뷰5", "리뷰6", "리뷰7", "리뷰8", "리뷰9", "리뷰10", "Q&A개수", "Q&A1", "Q&A2", "Q&A3", "Q&A4", "Q&A5", "Q&A6", "Q&A7", "Q&A8", "Q&A9", "Q&A10" ])
    for line in open('./resources/temp.txt'):
        if line != "":
            try:
                sheet.append([str(success), line.split('###')[1], line.split('###')[2], line.split('###')[3], line.split('###')[4], line.split('###')[5], line.split('###')[6], line.split('###')[7], line.split('###')[9], line.split('###')[10], line.split('###')[11], line.split('###')[12], line.split('###')[13], line.split('###')[14], line.split('###')[15], line.split('###')[16], line.split('###')[17], line.split('###')[18], line.split('###')[19], line.split('###')[20], line.split('###')[21], line.split('###')[22], line.split('###')[23], line.split('###')[24], line.split('###')[25], line.split('###')[26], line.split('###')[27], line.split('###')[28], line.split('###')[29], line.split('###')[30], line.split('###')[31], line.split('###')[32], line.split('###')[33]])
                success += 1
            except:
                fail += 1
                pass
    for cell in sheet["G"]:
        if cell.value == "링크":
            pass
        else:
            cell.hyperlink = cell.value
            cell.style = "Hyperlink"
    wb.save(f"./result.xlsx")
    msg = ctypes.windll.user32.MessageBoxW(None, f'결과 => 전체[{total}] 성공[{success}] 실패[{fail}]', "크롤링 완료!", 0)
    exit()