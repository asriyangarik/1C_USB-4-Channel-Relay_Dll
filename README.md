# 1C_USB-4-Channel-Relay_Dll

for use tis dll put the usb_relay_device.dll and LedControllCITY.dll in the bin folder of 1C for example C:\Program Files (x86)\1cv81\bin and registrated it by Regasm.exe open the comand line by Administrator and use this comand C:\Program Files (x86)\1cv81\bin\RegAsm.exe LedControllCITY.dll use RegAsm Microsoft.NET -> Framework -> v4.0.30319.

1C example

LEDControl = Новый COMОбъект("AddIn.LEDControl"); статус = LEDControl.MyDeviceConnect(); LEDControl.AllReleOff(); LEDControl.ReleOn(1); LEDControl.ReleOff(1); LEDControl.ReleOn(2); LEDControl.ReleOff(2);
