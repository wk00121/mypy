# -*- coding:UTF-8 -*-
#!/usr/bin/python
#-*- coding: utf-8 -*-
#encoding=utf-8
import urllib2
import urllib
import os
from BeautifulSoup import BeautifulSoup
def getAllImageLink():

    html = urllib2.urlopen('http://www.csdn.net/company/layer.html').read()
    soup = BeautifulSoup(html)
    liResult = soup.findAll('img',attrs={"alt":""})
    print liResult  #方便调试

    count = 0;
    for image in liResult:
        count += 1
        link = image.get('src')
        imageName = count
       # filesavepath = '/Users/justinjing/Desktop/testpython/%s.jpg' % imageName
        filesavepath = '/Users/Administrator/Desktop/test/%s.jpg' % imageName
        urllib.urlretrieve(link,filesavepath)
        print filesavepath
if __name__ == '__main__':
    getAllImageLink()