#!/usr/local/bin/python3.4
"""
第一个叫zs，第二个叫zhangs，第三个叫zhangsan，再多会加-
还有 s-zhang09
"""
import xlrd
import xlwt
import os
from unidecode import unidecode
from collections import namedtuple

os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 全名, 姓, 姓缩写, 名, 名缩写
Person = namedtuple('Person', ['fullname', 'xing', 'xing_s', 'ming', 'ming_s'])

class GenEmail():
    def __init__(self, filepath, year):
        self.input_xlspath = filepath
        self.year = year
        self.input_xls = xlrd.open_workbook(filepath)
        self.names_email = set()
        self.email_name_patterns = (  # don't know for sure, e.g.zhangsan
            lambda x: x.xing_s + x.ming_s,          # zs
            lambda x: x.xing + x.ming_s,            # zhangs
            lambda x: x.xing + x.ming,              # zhangsan
            lambda x: x.xing_s + '-' + x.ming_s,    # z-s
            lambda x: x.xing + '-' + x.ming_s,      # zhang-s
            lambda x: x.xing + '-' + x.ming,        # zhang-san
            lambda x: x.ming_s + x.xing,            # szhang
            lambda x: x.ming_s + '-' + x.xing,      # s-zhang
        )

    def parse_names(self):
        self.names = [self.input_xls.sheets()[0].cell(row, 0).value.strip() for row in range(1, self.input_xls.sheets()[0].nrows)]
        print('students: %d' % len(self.names))
        def make_person(name):
            xing = unidecode(name[0]).replace(' ', '').lower()
            xing_s = xing[0].lower()
            ming = unidecode(name[1:]).replace(' ', '').lower()
            ming_s = ''.join(unidecode(c)[0].lower() for c in name[1:])
            return Person._make([name, xing, xing_s, ming, ming_s])
        # ignore names with more than 3 charaters for now
        self.names_detail = tuple(make_person(name) for name in self.names if len(name) <= 3)
        print('students with name 2,3: %d' % len(self.names_detail))
        del self.names
        print(self.names_detail[:20])

    def make_email_name(self):
        def insert_to_names_email(name):
            if name not in self.names_email:
                self.names_email |= {name}
                return True
            else:
                return False

        for name in self.names_detail:
            for pattern in self.email_name_patterns:
                if insert_to_names_email(pattern(name)):
                    break

    def gen_email_address(self, email_name):
        return email_name + self.year + '@mails.tsinghua.edu.cn'


if __name__ == '__main__':
    gen_email = GenEmail('Fresh13.xls', '13')
    gen_email.parse_names()
    gen_email.make_email_name()
    with open('email13.txt', 'w') as f:
        for email_name in gen_email.names_email:
            f.write(gen_email.gen_email_address(email_name)+'\n')
