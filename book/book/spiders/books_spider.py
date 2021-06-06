import scrapy
import re
from unidecode import unidecode
from json_excel_converter import Converter
from json_excel_converter.xlsx import Writer


class BooksSpider(scrapy.Spider):
    name = 'books'
    crawled_books = []

    def start_requests(self):
        start_id = getattr(self, 'start', None)
        end_id = getattr(self, 'end', None)
        if (start_id is None) or (end_id is None):
            raise Exception('Please enter `start` and `end` correctly!')
        start_id = int(start_id)
        end_id = int(end_id)
        for id in range(start_id - 1, end_id + 1):
            url = f'https://db.ketab.ir/bookview.aspx?bookid={id}'
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        id = response.url.split('=')[1]
        title = response.css('span#ctl00_ContentPlaceHolder1_lblBookTitle::text').get()
        if title == 'Label':
            return
        subjects = response.xpath("//*[contains(@id, 'ctl00_ContentPlaceHolder1_rptSubject_ctl')]/text()").getall()
        authors = response.xpath("//*[contains(@id, 'ctl00_ContentPlaceHolder1_rptAuthor_ctl')]/text()").getall()
        publishers = response.xpath("//*[contains(@id, 'ctl00_ContentPlaceHolder1_rptPublisher_ctl')]/text()").getall()
        isbn = response.xpath("//*[@id='ctl00_ContentPlaceHolder1_lblISBN']/text()").getall()[1].strip()
        date = unidecode(response.xpath("//*[@id='ctl00_ContentPlaceHolder1_lblIssueDate']/text()").get())
        price = response.xpath("//*[@id='ctl00_ContentPlaceHolder1_lblprice']/strong/text()").get()
        if price is not None:
            price = unidecode(price)
        dioee = unidecode(response.xpath("//*[@id='ctl00_ContentPlaceHolder1_lblDoe']/a/text()").get())
        language = response.xpath("//*[@id='ctl00_ContentPlaceHolder1_Labellang']/a/text()").get()
        city = response.xpath("//*[@id='ctl00_ContentPlaceHolder1_lblplace']/text()").get()
        info = re.sub(' +', ' ', response.xpath("//*[@id='ctl00_ContentPlaceHolder1_lblISBN']/../text()").getall()[1].strip())
        description = response.xpath("//*[text()='معرفی مختصر كتاب']/../text()").getall()[2].strip()
        image = response.xpath("//*[@id='ctl00_ContentPlaceHolder1_imgBook']/@src").get()
        book = {
            'id': id,
            'title': title,
            'subjects': subjects,
            'authors': authors,
            'publishers': publishers,
            'isbn': isbn,
            'date': date,
            'price': price,
            'dioee': dioee,
            'language': language,
            'city': city,
            'info': info,
            'description': description,
            'image': image,
        }
        self.crawled_books.append(book)

    @classmethod
    def from_crawler(cls, crawler, *args, **kwargs):
        spider = super().from_crawler(crawler, *args, **kwargs)
        crawler.signals.connect(spider.spider_closed, signal=scrapy.signals.spider_closed)
        return spider

    def spider_closed(self, spider):
        conv = Converter()
        conv.convert(self.crawled_books, Writer(file=f"../{self.start}~{self.end}.xlsx"))
