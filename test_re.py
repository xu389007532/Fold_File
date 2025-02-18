import re
data='Cost-QNM034290-R1 4097773-2 SBA Easter 4 Cards （20241122 LL）.xlsx'
data='QNM033879-R2  4092633-2 IFG-STML Christmas Cards Acquisiton (240906 SF).xlsx'
re_QNM = re.compile('\S*QNM([0-9]{6})\S*', re.IGNORECASE)
re_QN = re.compile('\S*QN([0-9]{6})\S*', re.IGNORECASE)
print('QNM:',re_QNM.findall(data))
print('QN:',re_QN.findall(data))
QNM=re_QNM.findall(data)
if QNM:
    QNM_number=int(QNM[0])
    fold = 'QNM'+(str(int(QNM_number/1000))+'001').rjust(6,'0')+' to '+'QNM'+(str(int(33879/1000)+1)+'000').rjust(6,'0')
    print('Fold:',fold)