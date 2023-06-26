using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using USB_Relay_Control;

namespace LedControllCITY
{
    [ComVisible(true), Guid("0339E962-6744-4844-9408-1A795828CCBF"), ProgId("AddIn.LEDControl")]
    public class LEDControl : AddInLib.IInitDone, AddInLib.ILanguageExtender
    {
        const string c_AddinName = "LEDControl";

        string strMessageBoxIcon;
        string strMessageBoxButtons;
        int _deviceHandle = 0;


        private usb_relay_device_info _mydevice;


        #region "IInitDone implementation"

        public void Done()
        {

        }

        public void GetInfo(ref object pInfo)
        {
            ((System.Array)pInfo).SetValue("2000", 0);
        }

        public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            V7Data.V7Object = pConnection;
        }

        #endregion

        #region "Свойства"
        enum Props
        {   //Числовые идентификаторы свойств нашей внешней компоненты
            propMessageBoxIcon = 0,  //Пиктограмма для MessageBox'а
            propMessageBoxButtons = 1, //Кнопки для MessageBox'a
            LastProp = 2
        }

        public void GetNProps(ref int plProps)
        {   //Здесь 1С получает количество доступных из ВК свойств
            plProps = (int)Props.LastProp;
        }

        public void FindProp(string bstrPropName, ref int plPropNum)
        {   //Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
            switch (bstrPropName)
            {
                case "MessageBoxIcon":
                case "ПиктограммаПредупреждения":
                    plPropNum = (int)Props.propMessageBoxIcon;
                    break;
                case "MessageBoxButtons":
                case "КнопкиПредупреждения":
                    plPropNum = (int)Props.propMessageBoxButtons; break;
                default:
                    plPropNum = -1;
                    break;
            }
        }

        public void GetPropName(int lPropNum, int lPropAlias, ref string pbstrPropName)
        {   //Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
            pbstrPropName = "";
        }

        public void GetPropVal(int lPropNum, ref object pvarPropVal)
        {   //Здесь 1С узнает значения свойств 
            pvarPropVal = null;
            switch (lPropNum)
            {
                case (int)Props.propMessageBoxIcon:
                    pvarPropVal = strMessageBoxIcon;
                    break;
                case (int)Props.propMessageBoxButtons:
                    pvarPropVal = strMessageBoxButtons;
                    break;

            }
        }

        public void SetPropVal(int lPropNum, ref object varPropVal)
        {   //Здесь 1С изменяет значения свойств 

            switch (lPropNum)
            {
                case (int)Props.propMessageBoxIcon:
                    strMessageBoxIcon = (string)varPropVal;
                    break;
                case (int)Props.propMessageBoxButtons:
                    strMessageBoxButtons = (string)varPropVal;
                    break;
            }
        }

        public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
        {   //Здесь 1С узнает, какие свойства доступны для чтения
            pboolPropRead = true; // Все свойства доступны для чтения
        }

        public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
        {   //Здесь 1С узнает, какие свойства доступны для записи
            pboolPropWrite = true; // Все свойства доступны для записи
        }
        #endregion

        #region "Методы"

        enum Methods
        {   //Числовые идентификаторы методов (процедур или функций) нашей внешней компоненты
            methMessageBoxShow = 0, //Предупреждение, но с возможностью задавать пиктограмму и заголовок окна
            methExternalEvent = 1, //Генерирует внешнее событие (1С перехватывает его в процедуре ОбработкаВнешнегоСобытия())
            methShowErrorLog = 2, //Показываем ошибочное сообщение 
            methStatusLine = 3, //Показываем сообщение в строке состояния
            LastMethod = 4,
        }

        public void GetNMethods(ref int plMethods)
        {   //Здесь 1С получает количество доступных из ВК методов
            plMethods = (int)Methods.LastMethod;
        }

        public void FindMethod(string bstrMethodName, ref int plMethodNum)
        {   //Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции
            plMethodNum = -1;
            switch (bstrMethodName)
            {
                case "MessageBoxShow":
                case "Предупреждение":
                    plMethodNum = (int)Methods.methMessageBoxShow;
                    break;
                case "ExternalEvent":
                case "ВнешнееСобытие":
                    plMethodNum = (int)Methods.methExternalEvent;
                    break;
                case "ShowErrorLog":
                case "Сообщить":
                    plMethodNum = (int)Methods.methShowErrorLog;
                    break;
                case "StatusLine":
                case "Состояние":
                    plMethodNum = (int)Methods.methStatusLine;
                    break;

            }
        }

        public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
        {   //Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
            pbstrMethodName = "";
        }

        public void GetNParams(int lMethodNum, ref int plParams)
        {   //Здесь 1С получает количество параметров у метода (процедуры или функции)
            switch (lMethodNum)
            {
                case (int)Methods.methMessageBoxShow:
                    plParams = 2;
                    break;
                case (int)Methods.methExternalEvent:
                    plParams = 3;
                    break;
                case (int)Methods.methShowErrorLog:
                    plParams = 1;
                    break;
                case (int)Methods.methStatusLine:
                    plParams = 1;
                    break;
            }
        }

        public void GetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
        {   //Здесь 1С получает значения параметров процедуры или функции по умолчанию
            pvarParamDefValue = null; //Нет значений по умолчанию
        }

        public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
        {   //Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)
            pboolRetValue = true;  //Все методы у нас будут функциями (т.е. будут возвращать значение). 
        }

        public void CallAsProc(int lMethodNum, ref System.Array paParams)
        {   //Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.
        }

        public void CallAsFunc(int lMethodNum, ref object pvarRetValue, ref System.Array paParams)
        {   //Здесь внешняя компонента выполняет код функций.

            pvarRetValue = 0; //Возвращаемое значение метода для 1С

            switch (lMethodNum) //Порядковый номер метода
            {
                case (int)Methods.methMessageBoxShow: //Реализуем метод MessageBoxShow внешней компоненты
                    {

                        //Преобразовываем текстовое описание значка в MessageBoxIcon.ххх
                        switch (strMessageBoxIcon)
                        {
                            case "Asterisk":
                                //  icon = MessageBoxIcon.Asterisk;
                                break;
                            case "Error":
                                //   icon = MessageBoxIcon.Error;
                                break;
                            case "Exclamation":
                                //  icon = MessageBoxIcon.Exclamation;
                                break;
                            case "Hand":
                                //   icon = MessageBoxIcon.Hand;
                                break;
                            case "Information":
                                //   icon = MessageBoxIcon.Information;
                                break;
                            case "None":
                                //   icon = MessageBoxIcon.None;
                                break;
                            case "Question":
                                //   icon = MessageBoxIcon.Question;
                                break;
                            case "Stop":
                                //    icon = MessageBoxIcon.Stop;
                                break;
                            case "Warning":
                                //   icon = MessageBoxIcon.Warning;
                                break;
                        }


                        string strMessageBoxText;
                        string strMessageBoxHeader;
                        string strDialogResult;


                        strMessageBoxText = (string)paParams.GetValue(0); //Получаем первый параметр нашей функции - текст предупреждения
                        strMessageBoxHeader = (string)paParams.GetValue(1);//Получаем второй параметр нашей функции - заголовок предупреждения


                        strMessageBoxButtons = "OK";
                        strMessageBoxIcon = "None";
                        break;
                    } // конец метода MessageBosIcon
                      //////////////////////////////////////////////////////////

                case (int)Methods.methExternalEvent:  //Реализуем метод для генерации внешнего события
                    {
                        string s1;
                        string s2;
                        string s3;
                        s1 = (string)paParams.GetValue(0);
                        s2 = (string)paParams.GetValue(1);
                        s3 = (string)paParams.GetValue(2);
                        V7Data.AsyncEvent.ExternalEvent(s1, s2, s3);
                        break;
                    } // конец метода ExternalEvent
                      //////////////////////////////////////////////////////////

                case (int)Methods.methShowErrorLog:  //Реализуем метод для показа сообщения об ошибке
                    {
                        string s1;
                        s1 = (string)paParams.GetValue(0);

                        AddInLib.ExcepInfo ei = new AddInLib.ExcepInfo();
                        ei.wCode = 1006; //Вид пиктограммы
                        ei.bstrDescription = s1;
                        ei.bstrSource = c_AddinName;

                        V7Data.ErrorLog.AddError("", ei);
                        throw new System.Exception("An exception has occurred.");
                        //break;
                    } // конец метода ShowErrorLog
                      //////////////////////////////////////////////////////////

                case (int)Methods.methStatusLine: //Реализуем тестовый метод для изменения строки состояния
                    {
                        string s1;
                        s1 = (string)paParams.GetValue(0);
                        V7Data.StatusLine.SetStatusLine(s1);
                        System.Threading.Thread.Sleep(1000); //Делаем паузу 1 сек
                        V7Data.StatusLine.ResetStatusLine();
                        break;
                    }
            }
        }

        public void RegisterExtensionAs(ref string bstrExtensionName)
        {

        }
        #endregion

        #region MyMetods MyCode

        public string MyDeviceInfo()
        {
            try
            {
                usb_relay_device_info deviceInfo = RelayDeviceWrapper.usb_relay_device_enumerate();
                return deviceInfo.ToString();
            }
            catch (Exception ex)
            {

                return ex.Message;
            }


        }

        public string MyDeviceConnect()
        {
            try
            {
                usb_relay_device_info SelectedItem = RelayDeviceWrapper.usb_relay_device_enumerate();
                usb_relay_device_info device = SelectedItem;
                _mydevice = device;
                _deviceHandle = RelayDeviceWrapper.usb_relay_device_open(ref device);
                int numberOfRelays = (int)device.type;

                uint status = 0;
                RelayDeviceWrapper.usb_relay_device_get_status(_deviceHandle, ref status);

                return device.ToString() + "_Connected sucsesfull";
            }
            catch (Exception)
            {

                return "Cannot Connect To devise";
            }
          
        }

        public string[] MyDeviceStattus()
        {
            int numberOfRelays = (int)_mydevice.type;
            uint status = 0;
            RelayDeviceWrapper.usb_relay_device_get_status(_deviceHandle, ref status);

            string[] myRelesStatus = new string[numberOfRelays];
            for (int i = 0; i < numberOfRelays; i++)
            {
                
                if (status > numberOfRelays)
                    myRelesStatus[i] = $"{i}-On";
                else if (i + 1 == status)
                    myRelesStatus[i] = $"{i}-Off";
            }
            return myRelesStatus;
        }

        #region Rele On Off

        public void ReleOn(int ReleNum)
        {
            int openResult = RelayDeviceWrapper.usb_relay_device_open_one_relay_channel(_deviceHandle, ReleNum);
        }

        public void ReleOff(int ReleNum)
        {
            int Result = RelayDeviceWrapper.usb_relay_device_close_one_relay_channel(_deviceHandle, ReleNum);
        }

        public void AllReleOff()
        {
            for (int i = 0; i < Convert.ToInt16(_mydevice.type); i++)
            {
                int Result = RelayDeviceWrapper.usb_relay_device_close_one_relay_channel(_deviceHandle, i);
            }
        }

        #endregion

        #endregion


    }
}
