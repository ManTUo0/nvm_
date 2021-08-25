# encoding = utf-8

from bs4 import BeautifulSoup        # 网页解析，获取数据
import xlwt         # excel操作库，进行excel操作
import re           # 正则表达式，获取匹配内容
import urllib.request
import urllib.error


# 1.爬取网页
# 2.获取数据
# 3.保存数据
def main():
    baseURL = "https://movie.douban.com/top250?start="
    datalist = getData(baseURL)     # 爬取网页数据，返回列表
    savepath = "豆瓣电影Top250.xls"
    saveData(datalist, savepath)


# 正则表达式查找规则
# 影片详情链接
findLink = re.compile(r'<a href="(.*?)">')
# 影片封面链接
findImg = re.compile(r'<img.*src="(.*?)"', re.S)
# 片名
findTitle = re.compile(r'<span class="title">(.*)</span>')
# 影片分数
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
# 评价人数
findJudge = re.compile(r'<span>(\d*)人评价</span>')
# 引言
findInq = re.compile(r'<span class="inq">(.*)。</span>')
# 影片相关信息
findBd = re.compile(r'<p class="">(.*?)</p>', re.S)


# 爬取网页,逐一解析数据
def getData(baseURL):
    datalist = []
    for i in range(0, 10):
        url = baseURL + str(i*25)
        html = askURL(url)              # 保存获取到的源码
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all('div', class_="item"):        # find_all返回一个列表，每个列表元素为<div class="item">....</div>的源码，即每部电影的相关信息与详情链接所在的源码中
            # print(item)    # 测试是否找到源码
            data = []
            item = str(item)

            link = re.findall(findLink, item)[0]        # 找出影片详情链接(通常有两个，只用其中一个)
            data.append(link)

            imgSrc = re.findall(findImg, item)[0]       # 找出封面
            data.append(imgSrc)

            titles = re.findall(findTitle, item)        # 找出片名，中文外文
            if(len(titles) == 2):
                cn_title = titles[0]
                data.append(cn_title)                   # 加入中文名
                other_title = titles[1].replace('/', '')
                data.append(other_title)                # 加入外文名
            else:
                data.append(titles[0])
                data.append(' ')

            rating = re.findall(findRating, item)[0]    # 找出评分
            data.append(rating)

            judgeNum = re.findall(findJudge, item)      # 找出评论人数
            data.append(judgeNum)

            inq = re.findall(findInq, item)             # 找出引言
            if(len(inq) != 0):
                inq = inq[0]                            # 有引言存入
                data.append(inq)
            else:
                data.append(" ")                        # 无引言存入空格

            bd = re.findall(findBd, item)[0]               # 找出相关信息
            # 去掉一些无用格式
            bd = re.sub('<br(\s+)?/>(\s+)?', '', bd)
            bd = re.sub('/', '', bd)
            data.append(bd.strip())                     # 存入，且去掉前后空格

            datalist.append(data)
    return datalist


# 爬取网页源码
def askURL(baseURL):
    head = {
        "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36'
    }                           # 用户代理信息
    request = urllib.request.Request(baseURL, headers=head)         # 打包，模拟浏览器向网站发送请求
    html = ""
    try:
        response = urllib.request.urlopen(request)                  # 打开url链接
        html = response.read().decode('utf-8')                      # 读取返回数据，以utf-8的格式
        # print(html)
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


def saveData(datalist, savepath):
    print("Saving...")
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('豆瓣电影Top250', cell_overwrite_ok=True)
    col = ("电影详情链接", "图片链接", "影片中文名", "影片外国名", "评分", "评价数", "引言概况", "相关信息")
    for i in range(0, 8):
        sheet.write(0, i, col[i])               # 写入标识
    for i in range(0, 250):
        print("第%d条..." % (i+1))
        data = datalist[i]
        for j in range(0, 8):
            sheet.write(i+1, j, data[j])        # 存入电影数据
    book.save(savepath)
    print("爬取完毕！！")


if __name__ == "__main__":
    main()
