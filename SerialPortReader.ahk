#NoEnv
#SingleInstance force 

SetTitleMatchMode, 2
SendMode Input 
SetWorkingDir %A_ScriptDir%


ConfigFile := A_ScriptDir . "\" . ScriptName . ".ini"
TEXT_Exit := "Exit"

Menu, Tray, NoStandard
Menu, Tray, Add, %TEXT_Exit%, ExitProgram
Menu, Tray, Tip, %ScriptName%
 
 
IniRead, RS232_Port, %ConfigFile%, Section, COM_PORT
IniRead, RS232_Baud, %ConfigFile%,Section, COM_BAUD
IniRead, RS232_Parity, %ConfigFile%,Section, COM_PARITY
IniRead, RS232_Data, %ConfigFile%,Section, COM_DATA
IniRead, RS232_Stop, %ConfigFile%,Section, COM_STOP
IniRead, DTR_Status, %ConfigFile%,Section, DTR
IniRead, RTS_Status, %ConfigFile%,Section, RTS
IniRead, regex_string,%ConfigFile%,Section,REGEX
IniRead, var, %ConfigFile%,Section, SEND_VAR
IniRead,mode,%ConfigFile%,Section, MODE

if(mode)
mode_string=ASCII mode
else 
mode_string=HEX mode
 
MsgBox, %RS232_Port% `n%RS232_Baud%`n%RS232_Parity%`n%RS232_Data%`n%RS232_Stop%`n%DTR_Status%`n%RTS_Status%`n%regex_string%`n %mode_string%

RS232_Settings   = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=%DTR_Status% rts=%RTS_Status%

MsgBox, Begin RS232 %RS232_Port% reading
Quit_var = 0

RS232_FileHandle:=RS232_Initialize(RS232_Settings)

Loop
{
  Read_Data := RS232_Read(RS232_FileHandle,"0xFF",RS232_Bytes_Received)

  If (RS232_Bytes_Received > 0)
  {
   
    IF (mode)
    {
      ASCII =
      Read_Data_Num_Bytes := StrLen(Read_Data) / 2
      Loop %Read_Data_Num_Bytes%
      {
        StringLeft, Byte, Read_Data, 2
        StringTrimLeft, Read_Data, Read_Data, 2
        Byte = 0x%Byte%
        Byte := Byte + 0    
        ASCII_Chr := Chr(Byte)
        ASCII = %ASCII%%ASCII_Chr% 
      }
      SendRaw, %ASCII%
    }
    Else
      Send, %Read_Data%
    Critical, Off
  }
  if Quit_var = 1
    Break
}

RS232_Close(RS232_FileHandle)

MsgBox, AHK is now disconnected from %RS232_Port%
Return

RS232_Initialize(RS232_Settings)
{

  StringSplit, RS232_Temp, RS232_Settings, `:
  RS232_Temp1_Len := StrLen(RS232_Temp1)
  If (RS232_Temp1_Len > 4)   
    RS232_COM = \\.\%RS232_Temp1%
  Else
    RS232_COM = %RS232_Temp1%

  StringTrimLeft, RS232_Settings, RS232_Settings, RS232_Temp1_Len+1 

  VarSetCapacity(DCB, 28)
  BCD_Result := DllCall("BuildCommDCB"
       ,"str" , RS232_Settings
       ,"UInt", &DCB)
  If (BCD_Result <> 1)
  {
    MsgBox, There is a problem with Serial Port communication. `nFailed Dll BuildCommDCB, BCD_Result=%BCD_Result% `nThe Script Will Now Exit.
    Exit
    ExitApp
  }

  RS232_FileHandle := DllCall("CreateFile"
       ,"Str" , RS232_COM         
       ,"UInt", 0xC0000000
       ,"UInt", 3
       ,"UInt", 0
       ,"UInt", 3
       ,"UInt", 0
       ,"UInt", 0 
       ,"Cdecl Int")

  If (RS232_FileHandle < 1)
  {
    MsgBox, There is a problem with Serial Port communication. `nFailed Dll CreateFile, RS232_FileHandle=%RS232_FileHandle% `nThe Script Will Now Exit.
    Exit
    ExitApp

  }

  SCS_Result := DllCall("SetCommState"
       ,"UInt", RS232_FileHandle
       ,"UInt", &DCB)
  If (SCS_Result <> 1)
  {
    MsgBox, There is a problem with Serial Port communication. `nFailed Dll SetCommState, SCS_Result=%SCS_Result% `nThe Script Will Now Exit.
    RS232_Close(RS232_FileHandle)
    Exit
    ExitApp
  }

  ReadIntervalTimeout        = 0xffffffff
  ReadTotalTimeoutMultiplier = 0x00000000
  ReadTotalTimeoutConstant   = 0x00000000
  WriteTotalTimeoutMultiplier= 0x00000000
  WriteTotalTimeoutConstant  = 0x00000000

  VarSetCapacity(Data, 20, 0)
  NumPut(ReadIntervalTimeout,         Data,  0, "UInt")
  NumPut(ReadTotalTimeoutMultiplier,  Data,  4, "UInt")
  NumPut(ReadTotalTimeoutConstant,    Data,  8, "UInt")
  NumPut(WriteTotalTimeoutMultiplier, Data, 12, "UInt")
  NumPut(WriteTotalTimeoutConstant,   Data, 16, "UInt")

  SCT_result := DllCall("SetCommTimeouts"
     ,"UInt", RS232_FileHandle
     ,"UInt", &Data)
  If (SCT_result <> 1)
  {
    MsgBox, There is a problem with Serial Port communication. `nFailed Dll SetCommState, SCT_result=%SCT_result% `nThe Script Will Now Exit.
    RS232_Close(RS232_FileHandle)
    Exit
    ExitApp
    
  }
  
  Return %RS232_FileHandle%
}

RS232_Close(RS232_FileHandle)
{
  CH_result := DllCall("CloseHandle", "UInt", RS232_FileHandle)
  If (CH_result <> 1)
    MsgBox, Failed Dll CloseHandle CH_result=%CH_result%

  Return
}

RS232_Write(RS232_FileHandle,Message)
{
  SetFormat, Integer, DEC

  StringSplit, Byte, Message, `,
  Data_Length := Byte0

  VarSetCapacity(Data, Byte0, 0xFF)

  i=1
  Loop %Byte0%
  {
    NumPut(Byte%i%, Data, (i-1) , "UChar")
    i++
  }
  WF_Result := DllCall("WriteFile"
       ,"UInt" , RS232_FileHandle
       ,"UInt" , &Data
       ,"UInt" , Data_Length
       ,"UInt*", Bytes_Sent
       ,"Int"  , "NULL")
  If (WF_Result <> 1 or Bytes_Sent <> Data_Length)
    MsgBox, Failed Dll WriteFile to RS232 COM, result=%WF_Result% `nData Length=%Data_Length% `nBytes_Sent=%Bytes_Sent%
}

RS232_Read(RS232_FileHandle,Num_Bytes,ByRef RS232_Bytes_Received)
{
  SetFormat, Integer, HEX

  Data_Length  := VarSetCapacity(Data, Num_Bytes, 0x55)
  Read_Result := DllCall("ReadFile"
       ,"UInt" , RS232_FileHandle
       ,"Str"  , Data
       ,"Int"  , Num_Bytes
       ,"UInt*", RS232_Bytes_Received
       ,"Int"  , 0)

  If (Read_Result <> 1)
  {
    MsgBox, There is a problem with Serial Port communication. `nFailed Dll ReadFile on RS232 COM, result=%Read_Result% - The Script Will Now Exit.
    RS232_Close(RS232_FileHandle)
    Exit
    ExitApp
  }
  
  i = 0
  Data_HEX =
  Loop %RS232_Bytes_Received%
  {
    Data_HEX_Temp := NumGet(Data, i, "UChar")
    StringTrimLeft, Data_HEX_Temp, Data_HEX_Temp, 2
    Length := StrLen(Data_HEX_Temp)
    If (Length =1)
      Data_HEX_Temp = 0%Data_HEX_Temp%
    i++
    Data_HEX := Data_HEX . Data_HEX_Temp
  }

  SetFormat, Integer, DEC
  Data := Data_HEX

  Return Data
}


^F1::
Quit_var=1
return

::qqq::
RS232_Write(RS232_FileHandle,var) ;Send it out the RS232 COM port
return