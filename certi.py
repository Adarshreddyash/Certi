from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd 
import qrcode
import time 


def main():
    print("----- Certi-----")
    print("-----a Certificate Creator------")

    #Accept EXCEL file
    path="Input.xlsx"
    #print("Given path is "+path)

    #output direcotry path
    output_dir_path="/home/ash/Projects/Certi/Output/"

    #Accept Date From user
    date=input("Enter Date:")

    #open Excel File
    # To open Workbook 
    wb = xlrd.open_workbook(path) 
    sheet = wb.sheet_by_index(0)
    #ROWS and Columns
    R=sheet.nrows
    C=sheet.ncols

    basewidth = 100
    hsize=100;
    selectFont = ImageFont.truetype(r'/home/ash/Projects/Certi/fonts/Oswald/Oswald-Bold.ttf', 30)
    for i in range(1,R): 
       # print('|')
        comp=sheet.cell_value(i,0)
        name=sheet.cell_value(i,1)
        #take a image
        img = Image.open("purple_temp.jpg")
        draw = ImageDraw.Draw(img)
        draw.text( (435,483),"the "+comp+" competition", (0,0,0), font=selectFont)
        draw.text( (418,343),name, (0,0,0), font=selectFont)
        draw.text( (654,565),date, (0,0,0), font=selectFont)    
        
        #Make QR Code Of file Path
        QR_img = qrcode.make('https://adarshreddyash.github.io'+name+'.pdf')
        #wpercent = (basewidth / float(QR_img.size[0]))
        #hsize = int((float(QR_img.size[1]) * float(wpercent)))
        #print(hsize)
        QR_img = QR_img.resize((basewidth, hsize), Image.ANTIALIAS)
        #img.save('resized_image.jpg')

        #paste the QR_code on Our Image
        img.paste(QR_img,(309,570))
        #save changed Image
        img.save(output_dir_path+name+str(i)+".pdf", "PDF", resolution=100.0)
    
    print("Sucessfully created certificates!")
    input("Please Enter to Exit")

if __name__ == "__main__":
    main()
