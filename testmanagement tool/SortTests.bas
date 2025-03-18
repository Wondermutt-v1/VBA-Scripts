Attribute VB_Name = "SortTests"
Sub SortTests()

Sheets("Power Train Tests").Visible = True
Sheets("Chasis Tests").Visible = True
Sheets("Baler Tests").Visible = True
Sheets("Engine Tests").Visible = True
Sheets("Cotton Picker Specific").Visible = True
Sheets("Cab Tests").Visible = True
Sheets("Electrical Tests").Visible = True
Sheets("Hydraulic Tests").Visible = True
Sheets("Steering Systems").Visible = True
Sheets("Total Vehicle").Visible = True

Dim sourceSheet As Worksheet
Dim balerSheet As Worksheet
Dim powerTrainSheet As Worksheet
Dim engineSheet As Worksheet
Dim cottonSheet As Worksheet
Dim chasisSheet As Worksheet
Dim TMSheet As Worksheet
Dim elctrcSheet As Worksheet
Dim hydSheet As Worksheet
Dim steerSheet As Worksheet
Dim TotlVhcl As Worksheet
Dim sysName As Range
Dim ID As Range
Dim baleTRng As Range



Set destWB = ThisWorkbook


Set sourceSheet = Sheets("TR Data")
Set cottonSheet = Sheets("Cotton Picker Specific")
Set balerSheet = Sheets("Baler Tests")
Set engineSheet = Sheets("Engine Tests")
Set cabSheet = Sheets("Cab Tests")
Set chasisSheet = Sheets("Chasis Tests")
Set powerTrainSheet = Sheets("Power Train Tests")
Set elctrcSheet = Sheets("Electrical Tests")
Set hydSheet = Sheets("Hydraulic Tests")
Set steerSheet = Sheets("Steering Systems")
Set Brake = Sheets("Brake Tests")
Set Fuel = Sheets("Fuel Tests")
Set TotlVhcl = Sheets("Total Vehicle")

rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
'Set sysName = sourceSheet.Range(Cells(5, 9), Cells(rowCnt, 9))

'cotton As String
releaseStat = "No Longer Required"
cotton = "COTTON PICKER / HARVESTER SPECIFIC"
baler = "BALER SPECIFIC SYSTEMS"
engine = "ENGINE"
cab = "CAB"
chasis = "CHASSIS"
powtrn = "POWER TRAIN"
electric = "ELECTRICAL"
hydraulic = "HYDRAULIC SYSTEMS"
steering = "STEERING SYSTEM"
brk = "BRAKE SYSTEM"
fuelsys = "FUEL SYSTEM"
totalVhcl = "TOTAL VEHICLE"


'initialize counters
cottoncntr = 5
balercntr = 5
engineCntr = 5
cabCntr = 5
chasCntr = 5
pwrtrnCntr = 5
electCntr = 5
hydCntr = 5
steerCntr = 5
brkCntr = 5
tvhclCntr = 5
fuelCntr = 5

'MsgBox (rowCnt - 4)
sourceSheet.Rows(4).Copy balerSheet.Rows(4)
sourceSheet.Rows(4).Copy cottonSheet.Rows(4)
sourceSheet.Rows(4).Copy engineSheet.Rows(4)
sourceSheet.Rows(4).Copy cabSheet.Rows(4)
sourceSheet.Rows(4).Copy chasisSheet.Rows(4)
sourceSheet.Rows(4).Copy powerTrainSheet.Rows(4)
sourceSheet.Rows(4).Copy elctrcSheet.Rows(4)
sourceSheet.Rows(4).Copy hydSheet.Rows(4)
sourceSheet.Rows(4).Copy steerSheet.Rows(4)
sourceSheet.Rows(4).Copy Brake.Rows(4)
sourceSheet.Rows(4).Copy Fuel.Rows(4)
sourceSheet.Rows(4).Copy TotlVhcl.Rows(4)

balerSheet.Cells(1, 2).Value = "Last Updated:"
balerSheet.Cells(1, 3).Value = Now()
balerSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

cottonSheet.Cells(1, 2).Value = "Last Updated:"
cottonSheet.Cells(1, 3) = Now()
cottonSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

engineSheet.Cells(1, 2).Value = "Last Updated:"
engineSheet.Cells(1, 3) = Now()
engineSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

cabSheet.Cells(1, 2).Value = "Last Updated:"
cabSheet.Cells(1, 3) = Now()
cabSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

chasisSheet.Cells(1, 2).Value = "Last Updated:"
chasisSheet.Cells(1, 3) = Now()
chasisSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

powerTrainSheet.Cells(1, 2).Value = "Last Updated:"
powerTrainSheet.Cells(1, 3) = Now()
powerTrainSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

elctrcSheet.Cells(1, 2).Value = "Last Updated:"
elctrcSheet.Cells(1, 3) = Now()
elctrcSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

hydSheet.Cells(1, 2).Value = "Last Updated:"
hydSheet.Cells(1, 3) = Now()
hydSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

steerSheet.Cells(1, 2).Value = "Last Updated:"
steerSheet.Cells(1, 3) = Now()
steerSheet.Cells(1, 3).NumberFormat = "dd-mmm-yy"

Brake.Cells(1, 2).Value = "Last Updated:"
Brake.Cells(1, 3) = Now()
Brake.Cells(1, 3).NumberFormat = "dd-mmm-yy"

Fuel.Cells(1, 2).Value = "Last Updated:"
Fuel.Cells(1, 3) = Now()
Fuel.Cells(1, 3).NumberFormat = "dd-mmm-yy"

TotlVhcl.Cells(1, 2).Value = "Last Updated:"
TotlVhcl.Cells(1, 3) = Now()
TotlVhcl.Cells(1, 3).NumberFormat = "dd-mmm-yy"

'rowCnt = 30

'Clean for data import
engineSheet.Activate
engineSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
cottonSheet.Activate
cottonSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
balerSheet.Activate
balerSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
cabSheet.Activate
cabSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
chasisSheet.Activate
chasisSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
powerTrainSheet.Activate
powerTrainSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
elctrcSheet.Activate
elctrcSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
hydSheet.Activate
hydSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
steerSheet.Activate
steerSheet.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
Brake.Activate
Brake.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
Fuel.Activate
Fuel.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear
TotlVhcl.Activate
TotlVhcl.Range(Cells(5, 1), Cells(rowCnt, colCnt)).Clear






'Separate the tests into the separate systems spreadsheets
For i = 5 To rowCnt
    If sourceSheet.Cells(i, 7).Value <> releaseStat Then  'filters out the no longer required tests
    If sourceSheet.Cells(i, 7).Value <> "Closed" Then
    'If sourceSheet.Cells(i, 7).Value <> "Completed" Then
    
        If sourceSheet.Cells(i, 9).Value = cotton Then
            sourceSheet.Rows(i).Copy cottonSheet.Rows(cottoncntr)
            cottoncntr = cottoncntr + 1
        End If
     
        If sourceSheet.Cells(i, 9).Value = baler Then
            sourceSheet.Rows(i).Copy balerSheet.Rows(balercntr)
            balercntr = balercntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = engine Then
            sourceSheet.Rows(i).Copy engineSheet.Rows(engineCntr)
            engineCntr = engineCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = cab Then
            sourceSheet.Rows(i).Copy cabSheet.Rows(cabCntr)
            cabCntr = cabCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = chasis Then
            sourceSheet.Rows(i).Copy chasisSheet.Rows(chasCntr)
            chasCntr = chasCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = powtrn Then
            sourceSheet.Rows(i).Copy powerTrainSheet.Rows(pwrtrnCntr)
            pwrtrnCntr = pwrtrnCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = electric Then
            sourceSheet.Rows(i).Copy elctrcSheet.Rows(electCntr)
            electCntr = electCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = hydraulic Then
            sourceSheet.Rows(i).Copy hydSheet.Rows(hydCntr)
            hydCntr = hydCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = steering Then
            sourceSheet.Rows(i).Copy steerSheet.Rows(steerCntr)
            steerCntr = steerCntr + 1
        End If
        
        If sourceSheet.Cells(i, 9).Value = brk Then
            sourceSheet.Rows(i).Copy Brake.Rows(brkCntr)
            brkCntr = brkCntr + 1
        End If
        If sourceSheet.Cells(i, 9).Value = fuelsys Then
            sourceSheet.Rows(i).Copy Fuel.Rows(fuelCntr)
            fuelCntr = fuelCntr + 1
        End If
        If sourceSheet.Cells(i, 9).Value = totalVhcl Then
            sourceSheet.Rows(i).Copy TotlVhcl.Rows(tvhclCntr)
            tvhclCntr = tvhclCntr + 1
        End If
        
    'End If
    End If
    End If
   ' End If
Next

balercntr = balerSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(balercntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(balercntr, 27)).NumberFormat = "d-mmm-yy"

cabCntr = cabSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(cabCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(cabCntr, 27)).NumberFormat = "d-mmm-yy"

engineCntr = engineSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(engineCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(engineCntr, 27)).NumberFormat = "d-mmm-yy"

chasCntr = chasisSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(chasCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(chasCntr, 27)).NumberFormat = "d-mmm-yy"

pwrtrnCntr = powerTrainSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(pwrtrnCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(pwrtrnCntr, 27)).NumberFormat = "d-mmm-yy"

electCntr = elctrcSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(electCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(electCntr, 27)).NumberFormat = "d-mmm-yy"

hydCntr = hydSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(hydCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(hydCntr, 27)).NumberFormat = "d-mmm-yy"

steerCntr = steerSheet.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(steerCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(steerCntr, 27)).NumberFormat = "d-mmm-yy"

brkCntr = Brake.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(tvhclCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(tvhclCntr, 27)).NumberFormat = "d-mmm-yy"

fuelCntr = Fuel.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(tvhclCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(tvhclCntr, 27)).NumberFormat = "d-mmm-yy"

tvhclCntr = TotlVhcl.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(5, 18), Cells(tvhclCntr, 21)).NumberFormat = "d-mmm-yy"
Range(Cells(5, 26), Cells(tvhclCntr, 27)).NumberFormat = "d-mmm-yy"


Sheets("2024 planning").Activate
Sheets("Power Train Tests").Visible = False
Sheets("Chasis Tests").Visible = False
Sheets("Baler Tests").Visible = False
Sheets("Engine Tests").Visible = False
Sheets("Cotton Picker Specific").Visible = False
Sheets("Cab Tests").Visible = False
Sheets("Electrical Tests").Visible = False
Sheets("Hydraulic Tests").Visible = False
Sheets("Steering Systems").Visible = False
Sheets("Brake Tests").Visible = False
Sheets("Total Vehicle").Visible = False
End Sub

Sub HideTabs()
Sheets("2024 planning").Activate
Sheets("Power Train Tests").Visible = False
Sheets("Chasis Tests").Visible = False
Sheets("Baler Tests").Visible = False
Sheets("Engine Tests").Visible = False
Sheets("Cotton Picker Specific").Visible = False
Sheets("Cab Tests").Visible = False
Sheets("Electrical Tests").Visible = False
Sheets("Hydraulic Tests").Visible = False
Sheets("Steering Systems").Visible = False
Sheets("Brake Tests").Visible = False
Sheets("Total Vehicle").Visible = False
End Sub
