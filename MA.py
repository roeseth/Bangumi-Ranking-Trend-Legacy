#Coded by Roeseth
import urllib
import threading
from datetime import datetime
import bs4
from bs4 import BeautifulSoup
import xlwt
import sys
reload(sys)
sys.setdefaultencoding('gbk')
f = open('MAlist.txt', 'r')
print 'Reading List......'
itemlist = f.readlines()
print 'Done'
Workbook_overwrite_ok = True
style_Accu = xlwt.XFStyle()
style_Accu.num_format_str = 'YYYY-M-D h:mm'
style_inAccu = xlwt.XFStyle()
style_inAccu.num_format_str = 'YYYY-M-D'
for itemn in xrange(0, len(itemlist)):
        itemid = itemlist[itemn].strip('\n')
        wbook = xlwt.Workbook()
        sheet = wbook.add_sheet('id%s' % itemid )
        print '\nOpening item id:%s......' % itemid
        count = 1
        count_All = [0,0,0,0,0,0,0,0,0,0,0]
        count_Sum = 0
	db_Time = {}
	t={}
#The number of threads
	tsum = 4
	
        sheet.write(0, 1, 'ID')
        sheet.write(0, 2, itemid)
        sheet.write(0, 3, 'Average')
        sheet.write(0, 5, 'Valid Votes')
        sheet.write(1, 0, 'Marks')
        sheet.write(1, 1, 'Time(Accu)')
        sheet.write(1, 2, 'Average')
        sheet.write(2, 2, '=SUM(INDIRECT("A"&3):A3)/(ROW()-2)')
	sheet.write(1, 4, 'Time(inAccu)')
        sheet.write(1, 5, 'Count')

        url = 'http://bangumi.tv/subject/'+str(itemid)+'/collections'
        bgm = BeautifulSoup(urllib.urlopen(url))
        db_Marks = bgm.findAll('li', {'class' : 'user' or 'user odd'})
        bgm_title = bgm.title.string[3:]
        sheet.write(0, 0, bgm_title)
        
        def multiprc(pagepart, thread_num):
                global count
                global count_All
                global count_Sum
                global db_Time
                global wbook
                global sheet
                for pg in xrange(thread_num, 1000, tsum):
                        url = 'http://bangumi.tv/subject/'+str(itemid)+'/'+str(pagepart)+'?page='+str(pg)
                        bgm = BeautifulSoup(urllib.urlopen(url))
                        db_Marks = bgm.findAll('li', {'class' : 'user' or 'user odd'})
                        if len(db_Marks) == 0 :
                                break
                        print 'item %d out of %d : Scraping Page %d......' % (itemn+1, len(itemlist), pg)
                        for i in xrange(0, len(db_Marks)):
                                if len(db_Marks[i]('span')) == 1 :
                                        continue
                                mutex.acquire()
                                tmp_Str = int(db_Marks[i]('span')[1]['class'][0][5:])
                                sheet.write(count+1, 0, tmp_Str)
                                count_All[tmp_Str] += 1
                                count_Sum += tmp_Str
        			tmp_Time_inAccu = count+1, 1, db_Marks[i].p.string.split()[0]
                                tmp_Time_Accu = count+1, 1, db_Marks[i].p.string
                                if db_Time.has_key(str(tmp_Time_inAccu[2])) == False:
                                        db_Time[str(tmp_Time_inAccu[2])] = 1
                                else:
                        		db_Time[str(tmp_Time_inAccu[2])] += 1
                                sheet.write(count+1, 1, datetime.strptime(tmp_Time_Accu[2], '%Y-%m-%d %H:%M'), style_Accu)
                                count += 1
                                mutex.release()

        mutex = threading.Lock()
        
        print '\nitem %d out of %d : Collections Part:' % (itemn+1, len(itemlist))
        for tn in xrange(0, tsum):
                t[tn] = threading.Thread(target=multiprc, args=('collections', tn+1))
                t[tn].start()
        for tn in xrange(0, tsum):
                t[tn].join()
        
        print '\nitem %d out of %d : Doings Part:' % (itemn+1, len(itemlist))
        for tn in xrange(0, tsum):
                t[tn] = threading.Thread(target=multiprc, args=('doings', tn+1))
                t[tn].start()
        for tn in xrange(0, tsum):
                t[tn].join()
        
        print '\nitem %d out of %d : On Hold Part:' % (itemn+1, len(itemlist))
        for tn in xrange(0, tsum):
                t[tn] = threading.Thread(target=multiprc, args=('on_hold', tn+1))
                t[tn].start()
        for tn in xrange(0, tsum):
                t[tn].join()

        print '\nitem %d out of %d : Dropped Part:' % (itemn+1, len(itemlist))
        for tn in xrange(0, tsum):
                t[tn] = threading.Thread(target=multiprc, args=('dropped', tn+1))
                t[tn].start()
        for tn in xrange(0, tsum):
                t[tn].join()

        t_count = 1
	for key in db_Time:
                sheet.write(t_count+1, 4, datetime.strptime(key, '%Y-%m-%d'), style_inAccu)
                sheet.write(t_count+1, 5, db_Time[key])
                t_count += 1
	db_Time.clear()
	
        sheet.write(0, 4, float(count_Sum)/(count-1))
        sheet.write(0, 6, count-1)
        print 'Writing item id:%s has finished' % (itemid)
        wbook.save('id%s.xls' % (itemid))
        print 'id:%s Saved' % itemid
        del wbook
print 'All Files Saved'
f.close()
raw_input('Press Enter to Exit')
