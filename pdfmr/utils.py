from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

from glob import glob
from django.conf import settings
import os
import shutil
import openpyxl
import random, string
import time

from .custmize import merge_excel 

#pdfからテキスト情報を抽出する関数
def convert_pdf_to_txt(path):
    
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    laparams.detect_vertical = True # Trueにすることで綺麗にテキストを抽出できる
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    #openで対象のpdfを読み込む
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    maxpages = 0   #最大ページ数の指定
    fstr = ''
    for page in PDFPage.get_pages(fp, maxpages=maxpages):   #1ページ分の情報を取得する
        interpreter.process_page(page)   # process_page()で1ページ分の情報をテキストに変換

        str = retstr.getvalue()  #StringIO オブジェクト内に格納されているテキスト情報を取得する。
        fstr += str   #fstr変数に取得したテキスト情報を追記していく

    fp.close()
    device.close()
    retstr.close()
    return fstr

#Excelデータを生成する関数
def create_excel(upload_dir,user_name):
    #アップロードしたファイルの取り込み
    #PDFがアップロードされたディレクトリパス情報を指定
    upload_path = os.path.join(upload_dir, "*.pdf")    
    template_file = os.path.join(settings.MEDIA_ROOT, "template","請求書一覧ファイル.xlsx")
    #PDFから抽出したテキストデータの反映先となる日時付の請求書Excelファイルパスを生成
    timestr = time.strftime("%Y%m%d-%H%M%S")
    work_file =os.path.join(settings.MEDIA_ROOT, "temp", "請求書一覧ファイル_" + timestr + ".xlsx") 
    #最終的に出来上がった請求書Excelデータを配置するユーザのディレクトリパスを生成 
    user_dir = os.path.join(settings.MEDIA_ROOT , "excel", user_name) 
    #日時付のExcelファイルを生成
    file_list = glob(upload_path)   #uploadされたPDFファイルリストを取得
    shutil.copyfile(template_file, work_file)  #テンプレートファイルをコピー
    book = openpyxl.load_workbook(work_file)   #Excelファイルオープン

    result_list =[]
    for pdf in file_list:
        #アップロードされたPDFファイルを1つずつ読み込んでText化
        result_txt = convert_pdf_to_txt(pdf)
        result_list.append(result_txt)
    #Excelにデータをセット  
    err_message = merge_excel(book,result_list,work_file)
    if err_message:
        return err_message          
    #個人ディレクトリへコピー
    shutil.move(work_file, user_dir)
