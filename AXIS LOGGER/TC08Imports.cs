/**************************************************************************
*
* Filename:    TC08Imports.cs
*
* Copyright:   Pico Technology Limited 2011
*
* Author:      CPY
*
* Description:
*   This file contains all the .NET wrapper calls needed to support
*   the console example. It also has the enums and structs required
*   by the (wrapped) function calls.
*
* History:
*    23/05/2011 	CPY	Created
*
* Revision Info: "file %n date %f revision %v"
*						""
*
***************************************************************************/

using System;
using System.Runtime.InteropServices;
using System.Text;


public class Win32Interop
{
    [DllImport("crtdll.dll")]
    public static extern int _kbhit();
}


namespace TC08example
{
	unsafe class Imports
	{
		#region constants
		private const string _DRIVER_FILENAME = "usbtc08.dll";


		#endregion

        public AXIS_LOGGER.MAIN_FORM MAIN_FORM
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

		#region Driver enums


        public enum TempUnit : short 
        {   USBTC08_UNITS_CENTIGRADE, 
            USBTC08_UNITS_FAHRENHEIT,
            USBTC08_UNITS_KELVIN,
            USBTC08_UNITS_RANKINE
        }

		
		#endregion

		#region Driver Imports
		#region Callback delegates
		
		#endregion

		[DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_open_unit")]
		public static extern short TC08OpenUnit();

        [DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_close_unit")]
        public static extern short TC08CloseUnit(short handle);

        [DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_run")]
        public static extern short TC08Run(short handle,
                                           int interval
                                           );

        [DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_stop")]
        public static extern short TC08Stop(short handle);

        [DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_get_formatted_info")]
        public static extern short TC08GetFormattedInfo(short handle,
                                                        StringBuilder unit_info,
                                                        short string_length
                                                        );

        [DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_set_channel")]
        public static extern short TC08SetChannel(short handle,
                                                  short channel,
                                                  char tc_type
                                                  );

        [DllImport(_DRIVER_FILENAME, EntryPoint = "usb_tc08_get_single")]
        public static extern short TC08GetSingle(short handle,
                                                  float[] temp,
                                                  short *overflow_flags,
                                                  TempUnit units
                                                  );

	
		#endregion
	}
}

