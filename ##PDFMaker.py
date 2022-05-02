#!/usr/bin/env python3

import os
import fitz  # pip install PyMuPDF==1.19.6(此为其中一个可运行版本，并非强制)
# from fitz import open,Document
import re
from tqdm import tqdm


def reFilter(content: str, filter_re: str = r"") -> str:
    artical:str = None
    it = re.finditer(filter_re, content, re.S)
    for match in it: 
#         print (match.group(1), sep="\t", end="\n")
        artical = match.group()
    return artical

def isTabInside(toc:list):
    flag = True
    for i in toc:
        if '\t' not in i:#任何一个没有tab的话就不是全有
            flag = False
    # print(flag)
    return flag 

def pic2pdf(img_dir, filename):

    file_list = os.listdir(img_dir)
    if len(file_list) == 0:
        input(f"该文件夹为空文件夹 : {img_dir} 下没有文件!\n 按下任意键退出本程序")
    # file_list.sort()
    
    doc = fitz.open()#新建空白pdf文件
    
    multipleFlag = False
    if os.path.isdir(img_dir+os.sep+file_list[0]):
        multipleFlag = True
    
    # print(f"multipleFlag is :{multipleFlag}")
    
    tocFileName:str = '##TOC.txt'
    tocCopyFileName:str = '##TOCcopy.txt'

    if not multipleFlag:#非批量制作
        try:
            tocS = []
            if os.path.getsize(tocFileName) != 0:#如果TOC.txt文件有内容，则进入自定义目录制作
                TOC =  []
                filter = r'[0-9]+'
                with open(tocFileName,'r',encoding='utf-8') as tocFile:
                    TOC = tocFile.readlines()#读取每一行信息
                
                #去除末尾的转行
                for i in range(len(TOC)):
                    TOC[i] = TOC[i].strip('\n')

                if '\t' in TOC[0] : #若第一行有\t，那么至少不是情况2.已经保证至少有一行了TOC[0]不会超下标
                    if isTabInside(TOC) == True:#如果全都有\t，则是情况4，自定义目录名称
                        for content in TOC:
                            TOCstr = content.split('\t')
                            page = reFilter(TOCstr[0] if len(TOCstr)==2 else '1',filter)
                            name = TOCstr[1] if len(TOCstr)==2 else '1'
                            tocS.append([1, name, int(page), {'kind': 1, 'xref': 550, 'zoom': 0.0}])
                    else:#情况3，指定递增起点
                        TOCstr = TOC[0].split('\t')                       
                        star = reFilter(TOCstr[1] if len(TOCstr)==2 else '1',filter)

                        for index in range(len(TOC)):
                            TOCstr = TOC[index].split('\t')                             
                            page = reFilter(TOCstr[0] if len(TOCstr)>=1 else '1',filter)
                            tocS.append([1, f'chapter{index+int(star)}', int(page), {'kind': 1, 'xref': 550, 'zoom': 0.0}])
                else:#情况2，默认章节名1,2,……
                    for index in range(len(TOC)):
                        page = reFilter(TOC[index],filter)
                        tocS.append([1, f'chapter{index+1}', int(page if page != None else 1), {'kind': 1, 'xref': 550, 'zoom': 0.0}])




            for img in tqdm(file_list):
                imgdoc = fitz.open(img_dir+os.sep+img)  # 打开图片
                pdfbytes = imgdoc.convertToPDF()  # 使用图片创建单页的 PDF
                imgpdf = fitz.open("pdf", pdfbytes)
                doc.insertPDF(imgpdf)  # 将当前页插入文档
            # for i in tocS:
            #     print(i) 
            # return
            if len(tocS) != 0:
                doc.set_toc(tocS)
                os.system(f'copy /Y {tocFileName} {tocCopyFileName}')
                os.system(f'cd.> {tocFileName}')
                # print(f'copy /Y {tocFileName} {tocCopyFileName}')
                # print(f'cd.> {tocFileName}')
            doc.save(filename)  # 保存pdf文件
            print(f"\n\n>>>{filename} PDF文件制作完成!PDF保存路径：{filename} <<<")
        except ValueError as e1:
            print(f"可能是目录所指定的页数超过了实际最大页数，尝试调整TOC.txt内容。\n 错误信息{repr(e1)}")
        except BaseException as e:
            print(f"发生未知错误，请检查文件夹下只有图片，或只有子文件夹，二者不可混淆\n 错误信息{repr(e)}")
        finally:    
            doc.close()
            input("\n>>>按下任意键或右上角退出本程序<<<")
            return
    
    
    #批量制作
    try:
        print("进入批量制作模式>>>")
        tocS = []
        imgCount = 1
        chapterCount = 1
        for subdir in os.listdir(img_dir):#遍历每个子文件夹                  
            if os.path.isdir(img_dir+os.sep+subdir):#批量制作，遍历每个子文件夹下的文件
                #章节目录+1 
                tocS.append([1, f'chapter{chapterCount}', imgCount, {'kind': 1, 'xref': 550, 'zoom': 0.0}])
                # print(tocS)
                for img in tqdm(os.listdir(img_dir+os.sep+subdir)):
                    
                    #将图片转为PDF的一页
                    imgdoc = fitz.open(img_dir+os.sep+subdir+os.sep+img)  # 打开图片
                    pdfbytes = imgdoc.convertToPDF()  # 使用图片创建单页的 PDF
                    imgpdf = fitz.open("pdf", pdfbytes)
                    doc.insertPDF(imgpdf)  # 将当前页插入文档  
                                      
                    imgCount += 1               
                chapterCount += 1
                
        #保存章节、保存整个PDF
        doc.set_toc(tocS)
        doc.save(filename)  # 保存pdf文件
        print(f"\n\n>>>{filename} PDF文件制作完成!PDF保存路径：{filename} <<<")
    except BaseException as e:
        print(f"发生未知错误，请检查文件夹下只有图片，或只有子文件夹，二者不可混淆\n 错误信息{repr(e)}")
    finally:
        doc.close()
        input("\n>>>按下任意键或右上角退出本程序<<<")


if __name__ == '__main__':
    img_dir = input("请输入文件目录 :")
    if img_dir[0] == '\"':
        img_dir = img_dir[1:-1]#如果是带引号的形式，需要去掉引号
    filename = f"{img_dir}.pdf"
    pic2pdf(img_dir, filename)
