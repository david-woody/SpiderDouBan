# encoding: utf-8
import re
import sys
import urllib2

from  bs4 import BeautifulSoup

from douban import html_outputer

reload(sys)
sys.setdefaultencoding("utf-8")


# 设置cookie
# cookie_jar = cookielib.CookieJar()
# opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookie_jar))



class SpiderMain(object):
    def __init__(self):
        self.htmlOutPuter = html_outputer.HtmlOutputer()
        return

    def getData(self):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.63 Safari/537.36',
            'Accept': 'text / html, application / xhtml + xml, application / xml;q = 0.9, image / webp, * / *;q = 0.8',
            'Referer': 'https://book.douban.com/people/drunkdoggy/wish'
        }
        # 定义一个读取网页的方法
        url = "https://book.douban.com/people/70472267/collect"
        req = urllib2.Request(url, headers=headers)
        result = urllib2.urlopen(req).read()
        # print  result
        soup = BeautifulSoup(result, "html.parser")
        #找到用户ID
        imgs=soup.select("#db-usr-profile img")[0]
        username=imgs.get("alt")
        allTags = list()
        allBooks = {}
        # tags = soup.find_all("li", class_=" clearfix")
        # for tag in tags:
        #     print "标签:",tag.a.string,"数量:",tag.span.string,"链接:",tag.a.get("href")
        # 开始找所有的书籍
        tags = soup.find_all(class_="subject-item")
        for tag in tags:
            title = tag.find("h2").a.get("title")
            if tag.find("h2").find("span") is not None:
                subTitle = tag.find("h2").find("span").text
            else:
                subTitle = ""
            # print "Title=", title + subTitle
            # print tag
            tags = tag.find("span", class_="tags")
            if tags is None:
                continue
            tagStr = re.findall("标签\: (.*)", str(tags.text))[0]
            # print tagStr
            tagstrlist = tagStr.split(" ")
            for tag in tagstrlist:
                if allTags.__contains__(tag):
                    allBooks[tag].append(title + subTitle)
                    continue
                else:
                    allBooks[tag] = list()
                    allTags.append(tag)
                    allBooks[tag].append(title + subTitle)
        # 获取所有的页数链接
        # 找到页面中 “后页>”的位置  如果只有一页  页面中是找不到这个span的 则该用户读过的书只有一页 不超过15个
        nextTag = soup.find("span", class_="next")
        if nextTag is not None:
            pageCount = nextTag.previous_sibling.previous_sibling.string
        else:
            pageCount = 1
        footer = soup.find(class_="paginator");
        pageUrls = list()
        if nextTag is not None:
            hrefs = footer.find_all("a")
            for href in hrefs:
                if pageUrls.__contains__(href.get("href")):
                    continue
                pageUrls.append(href.get("href"))
        else:
            print "无爬取链接"
        for pageurl in pageUrls:
            req = urllib2.Request(pageurl, headers=headers)
            result = urllib2.urlopen(req).read()
            subSoup = BeautifulSoup(result, "html.parser")
            tags = subSoup.find_all(class_="subject-item")
            for tag in tags:
                title = tag.find("h2").a.get("title")
                if tag.find("h2").find("span") is not None:
                    subTitle = tag.find("h2").find("span").text
                else:
                    subTitle = ""
                # print "Title=", title + subTitle
                # print tag
                tags = tag.find("span", class_="tags")
                if tags is None:
                    continue
                tagStr = re.findall("标签\: (.*)", str(tags.text))[0]
                # print tagStr
                tagstrlist = tagStr.split(" ")
                for tag in tagstrlist:
                    if allTags.__contains__(tag):
                        allBooks[tag].append(title + subTitle)
                        continue
                    else:
                        allBooks[tag] = list()
                        allTags.append(tag)
                        allBooks[tag].append(title + subTitle)
        print 50 * " "
        print  "总共多少页:", pageCount
        # print "所有标签:", ','.join(allTags)
        # print "所有书:", allBooks
        print "标签总数:"
        return username,allTags, allBooks


if __name__ == "__main__":
    spider_object = SpiderMain()
    username,allTags, allBooks = spider_object.getData()
    spider_object.htmlOutPuter.writeDataDefault(username,allTags, allBooks)
    spider_object.htmlOutPuter.save("result.xls")
        # for book in allBooks[tag]:
        #     print 10 * "*" + book
