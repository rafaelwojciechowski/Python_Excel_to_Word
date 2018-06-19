# importowane biblioteki
import win32com.client as win32
import time
import sys,os
import datetime

# uruchamianie excela
excel = win32.gencache.EnsureDispatch("Excel.Application")
# sciezka do pliku excel
book = excel.Workbooks.Open(r"C:\Users\rwojciechowski\Desktop\Python AKK\Analiza ekonomiczno-finansowa_31.03.xlsx")
# nazwa arkusza w excelu
sheet1 = book.Worksheets(1)
sheet2 = book.Worksheets(2)
sheet3 = book.Worksheets(3)
sheet4 = book.Worksheets(4)
sheet5 = book.Worksheets(5)
# True - excel widoczny / False - excel ukryty
excel.Visible = False
time.sleep(1)
# zakres kopiowanej tablicy w cudzyslowiu
sheet1.Range("C304:F310").Copy()

# uruchamianie worda
word = win32.gencache.EnsureDispatch("Word.Application")
# sciezka do pliku word
doc = word.Documents.Open(r"C:\Users\rwojciechowski\Desktop\Python AKK\Analiza ekonomiczno-finansowa_31.03.doc")
wordpath = r"C:\Users\rwojciechowski\Desktop\Python AKK\Analiza ekonomiczno-finansowa_31.03.doc"
# True - word widoczny / False - word ukryty
word.Visible = False

# nazwa zakladki do ktorej ma sie skopiowac tabela
tabela1 = doc.Bookmarks("tabela_1").Range
tabela1.Paste()

# zamykanie excela bez zapisywania
#book.Close(False)
#excel.Quit()

#############################tabla 2############################################
#time.sleep(1)
sheet1.Range("B320:F328").Copy()
tabela2 = doc.Bookmarks("tabela_2").Range
tabela2.Paste()
#############################tabla 4_2015#######################################
#time.sleep(1)
sheet1.Range("E94:F133").Copy()
tabela4_2015_1 = doc.Bookmarks("tabela_4_2015_1").Range
tabela4_2015_1.Paste()
#time.sleep(1)
sheet1.Range("G94:H133").Copy()
tabela4_2015_2 = doc.Bookmarks("tabela_4_2015_2").Range
tabela4_2015_2.Paste()
#time.sleep(1)
sheet1.Range("I94:J133").Copy()
tabela4_2015_3 = doc.Bookmarks("tabela_4_2015_3").Range
tabela4_2015_3.Paste()
#############################tabla 4_2016#######################################
#time.sleep(1)
sheet1.Range("K94:L133").Copy()
tabela4_2016_1 = doc.Bookmarks("tabela_4_2016_1").Range
tabela4_2016_1.Paste()
#time.sleep(1)
sheet1.Range("M94:N133").Copy()
tabela4_2016_2 = doc.Bookmarks("tabela_4_2016_2").Range
tabela4_2016_2.Paste()
#time.sleep(1)
sheet1.Range("O94:P133").Copy()
tabela4_2016_3 = doc.Bookmarks("tabela_4_2016_3").Range
tabela4_2016_3.Paste()
#############################tabla 4_2017#######################################
#time.sleep(1)
sheet1.Range("Q94:R133").Copy()
tabela4_2017_1 = doc.Bookmarks("tabela_4_2017_1").Range
tabela4_2017_1.Paste()
#time.sleep(1)
sheet1.Range("S94:T133").Copy()
tabela4_2017_2 = doc.Bookmarks("tabela_4_2017_2").Range
tabela4_2017_2.Paste()
#time.sleep(1)
sheet1.Range("U94:V133").Copy()
tabela4_2017_3 = doc.Bookmarks("tabela_4_2017_3").Range
tabela4_2017_3.Paste()
#############################tabla 4_2018#######################################
#time.sleep(1)
sheet1.Range("AO94:AP133").Copy()
tabela4_2018_1 = doc.Bookmarks("tabela_4_2018_1").Range
tabela4_2018_1.Paste()
#time.sleep(1)
sheet1.Range("AQ94:AR133").Copy()
tabela4_2018_2 = doc.Bookmarks("tabela_4_2018_2").Range
tabela4_2018_2.Paste()
#time.sleep(1)
sheet1.Range("AS94:AT133").Copy()
tabela4_2018_3 = doc.Bookmarks("tabela_4_2018_3").Range
tabela4_2018_3.Paste()
#############################tabla 4_2022#######################################
#time.sleep(1)
sheet1.Range("AU94:AV133").Copy()
tabela4_2022_1 = doc.Bookmarks("tabela_4_2022_1").Range
tabela4_2022_1.Paste()
#time.sleep(1)
sheet1.Range("AW94:AX133").Copy()
tabela4_2022_2 = doc.Bookmarks("tabela_4_2022_2").Range
tabela4_2022_2.Paste()
#time.sleep(1)
sheet1.Range("AY94:AZ133").Copy()
tabela4_2022_3 = doc.Bookmarks("tabela_4_2022_3").Range
tabela4_2022_3.Paste()
#############################tabla 4_RAZEM#######################################
#time.sleep(1)
sheet1.Range("BA94:BB133").Copy()
tabela4_RAZEM_1 = doc.Bookmarks("tabela_4_RAZEM_1").Range
tabela4_RAZEM_1.Paste()
#time.sleep(1)
sheet1.Range("BC94:BD133").Copy()
tabela4_RAZEM_2 = doc.Bookmarks("tabela_4_RAZEM_2").Range
tabela4_RAZEM_2.Paste()
#time.sleep(1)
sheet1.Range("BE94:BF133").Copy()
tabela4_RAZEM_3 = doc.Bookmarks("tabela_4_RAZEM_3").Range
tabela4_RAZEM_3.Paste()
#############################tabla 5############################################
#time.sleep(1)
sheet4.Range("G373:G381").Copy()
tabela5 = doc.Bookmarks("tabela_5").Range
tabela5.Paste()
#############################tabla 6############################################
#time.sleep(1)
sheet4.Range("L8:L15").Copy()
tabela6_KW = doc.Bookmarks("tabela_6_KW").Range
tabela6_KW.Paste()
#time.sleep(1)
sheet4.Range("P8:P15").Copy()
tabela6_NKW = doc.Bookmarks("tabela_6_NKW").Range
tabela6_NKW.Paste()
#time.sleep(1)
sheet4.Range("M8:M15").Copy()
tabela6_VAT = doc.Bookmarks("tabela_6_VAT").Range
tabela6_VAT.Paste()
#time.sleep(1)
sheet4.Range("N8:N15").Copy()
tabela6_SUMA_BRUTTO = doc.Bookmarks("tabela_6_SUMA_BRUTTO").Range
tabela6_SUMA_BRUTTO.Paste()
#time.sleep(1)
SumaKW = sheet4.Range("L38").Formula = "=L8+L9+L10+L11+L12+L13+L14+L15+L16+L17+L18+L19+L20+L21+L22"
sheet4.Cells("38", "L").Copy()
tabela6_SUMA_KW = doc.Bookmarks("tabela_6_SUMA_KW").Range
tabela6_SUMA_KW.Paste()
SumaNKW = sheet4.Range("P38").Formula = "=P8+P9+P10+P11+P12+P13+P14+P15+P16+P17+P18+P19+P20+P21+P22"
sheet4.Cells("38", "P").Copy()
tabela6_SUMA_NKW = doc.Bookmarks("tabela_6_SUMA_NKW").Range
tabela6_SUMA_NKW.Paste()
SumaVAT = sheet4.Range("M38").Formula = "=M8+M9+M10+M11+M12+M13+M14+M15+M16+M17+M18+M19+M20+M21+M22"
sheet4.Cells("38", "M").Copy()
tabela6_SUMA_VAT = doc.Bookmarks("tabela_6_SUMA_VAT").Range
tabela6_SUMA_VAT.Paste()
Suma = sheet4.Range("N38").Formula = "=N8+N9+N10+N11+N12+N13+N14+N15+N16+N17+N18+N19+N20+N21+N22"
sheet4.Cells("38", "N").Copy()
tabela6_SUMA = doc.Bookmarks("tabela_6_SUMA").Range
tabela6_SUMA.Paste()
#############################tabla 7############################################
#time.sleep(1)
sheet4.Range("C44:F51").Copy()
tabela7_DANE = doc.Bookmarks("tabela_7_DANE").Range
tabela7_DANE.Paste()
#time.sleep(1)
SumaEFRR = sheet4.Range("C74").Formula = "=C44+C45+C46+C47+C48+C49+C50+C51+C52+C53+C54+C55+C56+C57"
sheet4.Cells("74", "C").Copy()
tabela7_EFRR = doc.Bookmarks("tabela_7_SUMA_EFRR").Range
tabela7_EFRR.Paste()
SumaZEW = sheet4.Range("D74").Formula = "=D44+D45+D46+D47+D48+D49+D50+D51+D52+D53+D54+D55+D56+D57"
sheet4.Cells("74", "D").Copy()
tabela7_SUMAZEW = doc.Bookmarks("tabela_7_SUMA_ZEW").Range
tabela7_SUMAZEW.Paste()
SumaWW = sheet4.Range("E74").Formula = "=E44+E45+E46+E47+E48+E49+E50+E51+E52+E53+E54+E55+E56+E57"
sheet4.Cells("74", "E").Copy()
tabela7_SUMAWW = doc.Bookmarks("tabela_7_SUMA_WW").Range
tabela7_SUMAWW.Paste()
Suma = sheet4.Range("F74").Formula = "=F44+F45+F46+F47+F48+F49+F50+F51+F52+F53+F54+F55+F56+F57"
sheet4.Cells("74", "F").Copy()
tabela7_SUMA = doc.Bookmarks("tabela_7_SUMA").Range
tabela7_SUMA.Paste()
#############################tabla 8############################################
#time.sleep(1)
sheet4.Range("C621:I650").Copy()
tabela8 = doc.Bookmarks("tabela_8").Range
tabela8.Paste()
#############################tabla 9############################################
#time.sleep(1)
sheet4.Range("D653").Copy()
tabela9_FNPV = doc.Bookmarks("tabela_9_FNPV").Range
tabela9_FNPV.Paste()
#time.sleep(1)
sheet4.Range("F653").Copy()
tabela9_FRR = doc.Bookmarks("tabela_9_FRR").Range
tabela9_FRR.Paste()
#############################tabla 10###########################################
#time.sleep(1)
sheet4.Range("C658:K687").Copy()
tabela10 = doc.Bookmarks("tabela_10").Range
tabela10.Paste()
#############################tabla 11###########################################
#time.sleep(2)
sheet4.Range("D690").Copy()
tabela11_FNPV = doc.Bookmarks("tabela_11_FNPV").Range
tabela11_FNPV.Paste()
#time.sleep(2)
sheet4.Range("F690").Copy()
tabela11_FRR = doc.Bookmarks("tabela_11_FRR").Range
tabela11_FRR.Paste()
#############################tabla 12###########################################
#time.sleep(1)
sheet4.Range("C1686:I1715").Copy()
tabela12 = doc.Bookmarks("tabela_12").Range
tabela12.Paste()
#############################tabla 13###########################################
#time.sleep(1)
sheet1.Range("C510:F539").Copy()
tabela13 = doc.Bookmarks("tabela_13").Range
tabela13.Paste()
#############################tabla 14###########################################
#time.sleep(1)
sheet1.Range("J510:M539").Copy()
tabela14 = doc.Bookmarks("tabela_14").Range
tabela14.Paste()
#############################tabla 15###########################################
#time.sleep(1)
sheet1.Range("Q510:T539").Copy()
tabela15 = doc.Bookmarks("tabela_15").Range
tabela15.Paste()
#############################tabla 16###########################################
#time.sleep(1)
sheet1.Range("C547:F576").Copy()
tabela16 = doc.Bookmarks("tabela_16").Range
tabela16.Paste()
#############################tabla 17###########################################
#time.sleep(1)
sheet1.Range("J547:M576").Copy()
tabela17 = doc.Bookmarks("tabela_17").Range
tabela17.Paste()
#############################tabla 18###########################################
#time.sleep(1)
sheet1.Range("C582:F611").Copy()
tabela18 = doc.Bookmarks("tabela_18").Range
tabela18.Paste()
#############################tabla 19###########################################
#time.sleep(1)
sheet1.Range("C617:F646").Copy()
tabela19 = doc.Bookmarks("tabela_19").Range
tabela19.Paste()
#############################tabla 20###########################################
#time.sleep(1)
sheet1.Range("C790:I819").Copy()
tabela20 = doc.Bookmarks("tabela_20").Range
tabela20.Paste()
#############################tabla 21###########################################
#time.sleep(1)
sheet1.Range("C862:D891").Copy()
tabela21 = doc.Bookmarks("tabela_21").Range
tabela21.Paste()
#############################tabla 22###########################################
#time.sleep(1)
sheet1.Range("C760:E760").Copy()
tabela22_GDY = doc.Bookmarks("tabela_22_GDY").Range
tabela22_GDY.Paste()
#time.sleep(1)
sheet1.Range("L760:N760").Copy()
tabela22_GDA = doc.Bookmarks("tabela_22_GDA").Range
tabela22_GDA.Paste()
#############################tabla 23###########################################
#time.sleep(1)
sheet1.Range("I765:L765").Copy()
tabela23 = doc.Bookmarks("tabela_23").Range
tabela23.Paste()
#############################tabla 24###########################################
#time.sleep(1)
sheet1.Range("C764:F764").Copy()
tabela24 = doc.Bookmarks("tabela_24").Range
tabela24.Paste()
#############################tabla 25###########################################
#time.sleep(1)
sheet1.Range("C826:E855").Copy()
tabela25 = doc.Bookmarks("tabela_25").Range
tabela25.Paste()
#############################tabla 26###########################################
#time.sleep(1)
sheet1.Range("C898:D927").Copy()
tabela26 = doc.Bookmarks("tabela_26").Range
tabela26.Paste()
#############################tabla 27###########################################
#time.sleep(1)
sheet5.Range("B30:C32").Copy()
tabela27_ENPV = doc.Bookmarks("tabela_27_ENPV").Range
tabela27_ENPV.Paste()
#time.sleep(1)
sheet5.Range("F30:F32").Copy()
tabela27_BC = doc.Bookmarks("tabela_27_BC").Range
tabela27_BC.Paste()
##########################zapisywanie i zamykanie###############################
fileName,ext = os.path.splitext(wordpath)
backupFileName = fileName + "-" + datetime.datetime.strftime(datetime.datetime.now(), "%d_%m_%Y") + "_SKRYPT" + ".doc"
savePath = os.path.join(os.path.dirname(wordpath),backupFileName)

doc.SaveAs(savePath)
doc.Close()
word.Quit()
