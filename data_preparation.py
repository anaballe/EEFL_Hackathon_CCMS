import xlrd 
loc = '/root/Downloads/Hackathon/Instant Raw Report (11).xlsx'
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
print("Lamp_Id,Timestamp,ActivePower,ApparentPower,Current_R,Current_Y,Current_B,R_Phase_kW,Y_Phase_kW,B_Phase_kW,Target")
for i in range(1,2677):
    Timestamp = sheet.cell_value(i,1)
    ActivePower = sheet.cell_value(i,2)
    Apparent_Power = sheet.cell_value(i,3)
    Current_R = sheet.cell_value(i,4)
    Current_Y = sheet.cell_value(i,5)
    Current_B =  sheet.cell_value(i,6)
    Load_Relay_R = sheet.cell_value(i,14)
    Load_Relay_Y = sheet.cell_value(i,15)
    Load_Relay_B = sheet.cell_value(i,16)
    R_Phase_kW = sheet.cell_value(i,17)
    Y_Phase_kW = sheet.cell_value(i,18)
    B_Phase_kW = sheet.cell_value(i,19)
    if(Load_Relay_R == 'DISCONNECTED' or Load_Relay_Y == 'DISCONNECTED' or Load_Relay_B == 'DISCONNECTED'):
        Target = 1
    else :
        Target = 0
    t = Timestamp[11:13]
    tt = int(t)
    k = i % 40
    if((tt >= 17) or (tt <=5)):          print(k,",",Timestamp,",",ActivePower,",",Apparent_Power,",",Current_R,",",Current_Y,",",Current_B,",",R_Phase_kW,",",Y_Phase_kW,",",B_Phase_kW,",",Target)
    
    
