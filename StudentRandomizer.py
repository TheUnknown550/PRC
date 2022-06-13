import random as rn
from openpyxl import Workbook, load_workbook


#ใส่ชื่อคนห้องสายวิทย์
SciM5 = ['Matt','FookFick','Hana','Able','Chris']
SciM4 = ['Bibua','Yok','beebell','Kan','Mos','John','Sam']
Scipair5 = []
Scipair4 = []

#ใส่ชื่อคนห้องสานศิล
ArtM4 = ['Matt','FookFick','Hana','Able','Chris']
ArtM5 = ['Bibua','Yok','beebell','Kan','Mos','John','Sam']
Artpair5 = []
Artpair4 = []

def Sci():
    for i in range(len(SciM5)):
        rand = rn.randint(0,len(SciM4)-1)

        Scipair4.append(SciM4[rand])
        Scipair5.append(SciM5[i])
        SciM4.remove(SciM4[rand])

def Art():
    #Remove Team
    ArtM5.remove('Yok')
    ArtM5.remove('John')
    #ArtM5.remove('Yok')
    #ArtM5.remove('Yok')
    #ArtM5.remove('Yok')

    for i in range(len(ArtM4)):
        rand = rn.randint(0,len(ArtM5)-1)

        Artpair5.append(ArtM5[rand])
        Artpair4.append(ArtM4[i])
        ArtM5.remove(ArtM5[rand])
        print(Artpair5[i],'+',Artpair4[i])
        print(ArtM5)

if __name__ == "__main__":
    #Access Excel
    wb = Workbook()
    ws = wb.active

    #Randomizer
    Sci()
    Art()

    #Sci
    ws['A2'].value = "SCI"
    ws['A3'].value = "M5"
    ws['B3'].value = "M4"
    for i in range(len(Scipair5)):
        ws['A'+str(i+4)].value = Scipair5[i]
        ws['B'+str(i+4)].value = Scipair4[i]
        #Change Group Team
        if Scipair5[i] == 'Matt':
            rand = rn.randint(0,len(SciM4)-1)
            ws['C'+str(i+4)].value = SciM4[i]

    #Art
    ws['A'+str(i+7)].value = 'Art'
    ws['A'+str(i+8)].value = "M4"
    ws['B'+str(i+8)].value = "M5"
    for n in range(len(Artpair5)):
        ws['B'+str(i+9+n)].value = Artpair5[n]
        ws['A'+str(i+9+n)].value = Artpair4[n]
        
        #Change Group team
        if Artpair5[n] == 'Mos':
            ws['C'+str(i+9+n)].value = 'Yok'

        if Artpair5[n] == 'Sam':
            ws['C'+str(i+9+n)].value = 'John'

        if Artpair5[n] == 'Mos':
            ws['C'+str(i+9+n)].value = 'Yok'

        if Artpair5[n] == 'Mos':
            ws['C'+str(i+9+n)].value = 'Yok'

        if Artpair5[n] == 'Mos':
            ws['C'+str(i+9+n)].value = 'Yok'
            

    wb.save('C:/Users/mattc/OneDrive/Documents/PRC/StudentPair.xlsx')
