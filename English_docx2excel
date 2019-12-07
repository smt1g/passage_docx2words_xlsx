
import pandas
import docx
from tkinter import _flatten
import os
import openpyxl

def get_file(path,file_list):
    for file in os.listdir(path):
        if file.endswith(".docx"):
            file_list.append(file)
    #print(file_list)

def paragraphs_list(file_name):
    doc = docx.Document(file_name)
    paragraphs = []
    for i in range(len(doc.paragraphs)):
        p = doc.paragraphs[i]
        if len(p.text) > 10:
            paragraphs.append(p.text)
    return paragraphs



def para2words(paragraphs_list):
    word_list = []
    for i in range(len(paragraphs_list)):
        line=str(paragraphs_list.pop())
        line=filter(line)#过滤每一段
        word=line.split(' ')
        word_list.append(word)
    word_list=list(_flatten(word_list))#把多为数组转化为一维数组
    return pandas.value_counts(word_list) #return word:count

def filter(line): #处理掉,？（）这种东西
    filter_list=[',','?',':',')','(','.','，','）','（'] #要替换的符号
    for i in filter_list:
        line=line.replace(i,' ')
    line.strip()
    #print(line)
    return line
def save2excel():
    pass


def main():
    path = "D:\\python_project\\doc"
    file_list = []
    get_file(path,file_list)
    for file_name in file_list:
        print("***************"+file_name)
        paragraphs=paragraphs_list(file_name)
        word_list=para2words(paragraphs)
    word_list=list(word_list.items())
    #print(word_list[100][1])
    wb=openpyxl.Workbook()
    ws=wb.active
    count=1
    for i in range(1,50):
        for j in  range(1,len(word_list)//50,2):
            ws.cell(row=i,column=j).value=word_list[count][0]
            #print(word_list[count][0])
            ws.cell(row=i,column=j+1).value=word_list[count][1]
            #print(word_list[count][1])

            count=count+1



    # for i in range(len(word_list)):
    #     ws.append([word_list[i][0],word_list[i][1]])
    wb.save('words.xlsx')
main()
