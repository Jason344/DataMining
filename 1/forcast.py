
# -*- coding: utf-8 -*-
import xlrd
#引用外部模块处理Excel文件，安装python后，命令行输入easy_install xlrd
num = 34#号码区为1-33
def getDatabase():
    data = xlrd.open_workbook('datebase.xlsx')
    table = data.sheets()[0] 
    nrows = table.nrows
    start = 2 #第一个红球所在列
    end = 8   #最后一个红球所在列
    D = []
    for i in range(nrows):
        s = set()
        if i != 0: #第一行为表头，不用处理 
            for j in range(start,end):
                s.add(int(table.cell(i,j).value))
        D.append(s)
    return D

#根据集合列表返回每一个集合在数据库中出现的次数                  
def getDBcount(D,S):
    count = [0]*len(S)
    for record in D:
        T = getsubsWithSize(record,len(S[0]))
        for i in range(len(S)):
            if S[i] in T:#判断候选项是不是当前记录的子集
                count[i]=count[i]+1
    return count
#获取固定大小的子集，不用细看，我也是抄别人的
def getsubsWithSize(father, size):
    all_subs = []
    excludes = set()
    for i in father:
        sub = set()
        sub.add(i)
        if len(sub) == size:
            all_subs.append(sub)
        else:
            excludes.add(i)
            new_father = father.difference(excludes)
            _sub = getsubsWithSize(new_father, size-1)
            for s in _sub:
                all_subs.append(sub.union(s))
    return all_subs
#判断c有没有不频繁子集，同书上伪码
def hasInfreSubset(c,L):
    c_len = len(c)
    subsets = getsubsWithSize(c,c_len-1)
    for item in subsets:
        if(item not in L):
            return True
    return False
#生成下一代候选频繁项
def apriori_gen(L):
    C = []
    #把L中的项，两两匹配
    for item1 in range(len(L)):
        for item2 in range(item1+1,len(L)):
            c = set()
            s1 = L[item1]
            s2 = L[item2]
            size = len(s1)
            if(len(s1&s2)==size-1):#判断两项是不是只有一个不同
                c = s1|s2 #合并生成新的项，大小增加了1
                if(c not in C and not hasInfreSubset(c,L)):
                    C.append(c)
    return C
#找大小为1的频繁项
def findFreItems1(D,minSup):
    L=[]
    count = [0]*num
    #遍历每条记录的每个项目，并记录每个号的球出现了几次
    for record in D:
        for item in record:
            count[item]= count[item]+1
    #将频繁的项放入L中
    for item in range(len(count)):
        if(count[item]>=minSup):
            temp = set()
            temp.add(item)
            L.append(temp)
    return L

def apriori(D,minSup):
    L = findFreItems1(D,minSup)
    while len(L)!=0:
        C = apriori_gen(L)
        if(len(C)==0):
            break
        # 每一个统计候选项出现的频率
        count = getDBcount(D,C)
        # 将符合条件的候选项作为真正的频繁项
        L = []
        for i in range(len(C)):
            if(count[i]>=minSup):
                L.append(C[i])
    return L
#把频繁项按出现次数排序，并返回出现次数最多的项
def getMostFre(D,S):
    count = getDBcount(part,S)
    combine = zip(S,count)
    combine.sort(key=lambda x:(x[1]),reverse=True)
    return combine[0][0]
D = getDatabase()
A =[set([13, 22])]#apriori(D,30)
B =[set([1, 2, 22, 26]), set([1, 3, 8, 31]), set([1, 3, 21, 31]), set([32, 1, 3, 31]), set([1, 19, 5, 27]),
 set([1, 19, 7, 27]), set([1, 22, 28, 30]), set([2, 8, 29, 31]), set([2, 19, 27, 12]), set([17, 2, 20, 13]),
 set([2, 24, 29, 15]), set([3, 6, 11, 29]), set([4, 21, 7, 26]), set([19, 4, 22, 9]), set([17, 4, 12, 30]),
 set([19, 4, 25, 12]), set([19, 4, 22, 25]), set([5, 6, 8, 22]), set([5, 6, 12, 14]), set([18, 5, 6, 13]),
 set([5, 22, 13, 30]), set([16, 5, 30, 14]), set([32, 5, 22, 26]), set([19, 6, 28, 13]), set([17, 6, 14, 15]),
 set([6, 25, 14, 15]), set([17, 22, 6, 14]), set([16, 17, 6, 23]), set([7, 23, 14, 31]), set([7, 25, 26, 14]),
 set([24, 21, 8, 11]), set([27, 21, 8, 11]), set([22, 8, 29, 15]), set([19, 9, 10, 15]), set([23, 24, 9, 29]),
 set([25, 24, 9, 26]), set([32, 10, 11, 12]), set([32, 10, 11, 15]), set([19, 10, 29, 14]), set([20, 23, 10, 15]),
 set([18, 20, 27, 13]), set([19, 22, 26, 13]), set([17, 21, 25, 14]), set([17, 27, 28, 14]), set([16, 17, 24, 25]), set([18, 20, 28, 29])]#apriori(D,3)

part = D[-100:]
ans = getMostFre(part,B)|A[0]
l=list(ans)
l.sort()
print l
