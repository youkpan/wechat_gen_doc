#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import re
import hashlib
import html

if sys.version_info[0] == 2:
    from urllib2 import urlopen  # Python 2
else:
    from urllib.request import urlopen  # Python3

def save_video(videopath,url):
    data = urlopen(url).read()
    print("get data finish ",videopath,url)
    vfile = open(videopath, 'wb')
    vfile.write(data)
    vfile.close()

filename = "search"
count=0
vfile = open('./html/%s.html'%filename, 'rb')
html_str = ""
for line in vfile.readlines():
    html_str += line.decode("utf-8")

'''
matchObj = re.search( r'<source src="([^"]*)"', html_str, re.M|re.I)
print("len(matchObj.group(1))",len(matchObj.group(1)))
url = html.unescape(matchObj.group(1))
if matchObj:
   print ("search --> searchObj.group(1) : ", urlopen)
'''

pattern = re.compile( r'<source src="([^"]*)"') 
result1 = pattern.findall(html_str)
print(result1)

video_url_list = result1
#video_url_list = []
for url in video_url_list:
    try:
        url1 = html.unescape(url)

        print(count,count/len(video_url_list)*100,"%%, download: ",url1)
        hl = hashlib.md5()

        hl.update(url1.encode(encoding='utf-8'))

        videofilename = hl.hexdigest()
        videopath= './html/%s_files/%s.mp4'%(filename,videofilename)
        #print(videopath,'videofilename ：' + hl.hexdigest())

        html_str = html_str.replace("<source src=\""+url,"<source src=\"./"+filename+"_files/"+videofilename+".mp4",1)

        if os.path.exists(videopath): 
            print("exists",url1)
            continue

        save_video(videopath,url1)


        count +=1
        print('download count',count)

    except Exception as e:
        
        raise e

vfile.close()

vfile2 = open('./html/%s_ok.html'%filename, 'wb')
vfile2.write(html_str.encode('utf-8'))

vfile2.close()