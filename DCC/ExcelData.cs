using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToInterface
{
    class ExcelData
    {
        public ArrayList ArrayFrekansSParam { get; set; }
        public ArrayList ArrayFrekansEE { get; set; }
        public ArrayList ArrayFrekansCF { get; set; }
        public ArrayList ArrayFrekansCIS { get; set; }
        public ArrayList ArrayFrekansNoise { get; set; }

        // S11
        public ArrayList ArrayS11Reel { get; set; }
        public ArrayList ArrayS11ReelUnc { get; set; }
        public ArrayList ArrayS11Complex { get; set; }
        public ArrayList ArrayS11ComplexUnc { get; set; }
        public ArrayList ArrayS11Lin { get; set; }
        public ArrayList ArrayS11LinUnc { get; set; }
        public ArrayList ArrayS11LinPhase { get; set; }
        public ArrayList ArrayS11LinPhaseUnc { get; set; }
        public ArrayList ArrayS11Log { get; set; }
        public ArrayList ArrayS11LogUnc { get; set; }
        public ArrayList ArrayS11LogPhase { get; set; }
        public ArrayList ArrayS11LogPhaseUnc { get; set; }
        public ArrayList ArrayS11SWR { get; set; }
        public ArrayList ArrayS11SWRUnc { get; set; }

        // S12
        public ArrayList ArrayS12Reel { get; set; }
        public ArrayList ArrayS12ReelUnc { get; set; }
        public ArrayList ArrayS12Complex { get; set; }
        public ArrayList ArrayS12ComplexUnc { get; set; }
        public ArrayList ArrayS12Lin { get; set; }
        public ArrayList ArrayS12LinUnc { get; set; }
        public ArrayList ArrayS12LinPhase { get; set; }
        public ArrayList ArrayS12LinPhaseUnc { get; set; }
        public ArrayList ArrayS12Log { get; set; }
        public ArrayList ArrayS12LogUnc { get; set; }
        public ArrayList ArrayS12LogPhase { get; set; }
        public ArrayList ArrayS12LogPhaseUnc { get; set; }

        // S21
        public ArrayList ArrayS21Reel { get; set; }
        public ArrayList ArrayS21ReelUnc { get; set; }
        public ArrayList ArrayS21Complex { get; set; }
        public ArrayList ArrayS21ComplexUnc { get; set; }
        public ArrayList ArrayS21Lin { get; set; }
        public ArrayList ArrayS21LinUnc { get; set; }
        public ArrayList ArrayS21LinPhase { get; set; }
        public ArrayList ArrayS21LinPhaseUnc { get; set; }
        public ArrayList ArrayS21Log { get; set; }
        public ArrayList ArrayS21LogUnc { get; set; }
        public ArrayList ArrayS21LogPhase { get; set; }
        public ArrayList ArrayS21LogPhaseUnc { get; set; }

        // S22
        public ArrayList ArrayS22Reel { get; set; }
        public ArrayList ArrayS22ReelUnc { get; set; }
        public ArrayList ArrayS22Complex { get; set; }
        public ArrayList ArrayS22ComplexUnc { get; set; }
        public ArrayList ArrayS22Lin { get; set; }
        public ArrayList ArrayS22LinUnc { get; set; }
        public ArrayList ArrayS22LinPhase { get; set; }
        public ArrayList ArrayS22LinPhaseUnc { get; set; }
        public ArrayList ArrayS22Log { get; set; }
        public ArrayList ArrayS22LogUnc { get; set; }
        public ArrayList ArrayS22LogPhase { get; set; }
        public ArrayList ArrayS22LogPhaseUnc { get; set; }
        public ArrayList ArrayS22SWR { get; set; }
        public ArrayList ArrayS22SWRUnc { get; set; }


        // Effective 
        public ArrayList ArrayEffiencyEEEE { get; set; }
        public ArrayList ArrayEffiencyEEEEUnc { get; set; }
        public ArrayList ArrayEffiencyEE_S11Reel { get; set; }
        public ArrayList ArrayEffiencyEE_S11ReelUnc { get; set; }
        public ArrayList ArrayEffiencyEE_S11Imag { get; set; }
        public ArrayList ArrayEffiencyEE_S11ImagUnc { get; set; }
        public ArrayList ArrayEffiencyRHO_EERho { get; set; }
        public ArrayList ArrayEffiencyRHO_EERhoUnc { get; set; }
        public ArrayList ArrayEffiencyEE_CFEE_CF { get; set; }
        public ArrayList ArrayEffiencyEE_CFEE_CFUnc { get; set; }

        //Cf-Factor
        public ArrayList ArrayCF_Cal_Factor { get; set; }
        public ArrayList ArrayCF_Cal_Factor_Unc { get; set; }
        public ArrayList ArrayCF_Reel { get; set; }
        public ArrayList ArrayCF_Reel_Unc { get; set; }
        public ArrayList ArrayCF_Imaginer { get; set; }
        public ArrayList ArrayCF_Imaginer_Unc { get; set; }
        public ArrayList ArrayCF_ReflectionCof { get; set; }
        public ArrayList ArrayCF_ReflectionCof_Unc { get; set; }

        //CIS 
        public ArrayList ArrayCIS_Z_Position { get; set; }
        public ArrayList ArrayCIS_Z_Position_Unc { get; set; }
        public ArrayList ArrayCIS_ICOD { get; set; }
        public ArrayList ArrayCIS_ICOD_Unc { get; set; }
        public ArrayList ArrayCIS_OCID { get; set; }
        public ArrayList ArrayCIS_OCID_Unc { get; set; }


        //Noise
        public ArrayList ArrayNoiseENR { get; set; }
        public ArrayList ArrayNoiseENRUnc { get; set; }
        public ArrayList ArrayNoiseDCONRCLin { get; set; }
        public ArrayList ArrayNoiseDCONUpLimit { get; set; }
        public ArrayList ArrayNoiseDCONRCLinUnc { get; set; }
        public ArrayList ArrayNoiseDCONRCPhase { get; set; }
        public ArrayList ArrayNoiseDCONRCPhaseUnc { get; set; }
        public ArrayList ArrayNoiseDCONControl { get; set; }
        public ArrayList ArrayNoiseDCOFFRCLin { get; set; }
        public ArrayList ArrayDCOFFUpLimit { get; set; }
        public ArrayList ArrayNoiseDCOFFRCLinUnc { get; set; }
        public ArrayList ArrayNoiseDCOFFRCPhase { get; set; }
        public ArrayList ArrayNoiseDCOFFRCPhaseUnc { get; set; }
        public ArrayList ArrayNoiseDCOFFControl { get; set; }





        // Device Informations

        public string OrderNumber { get; set; }
        public string DeviceName { get; set; }
        public string SerialNumber { get; set; }


        public ExcelData()
        {
            ArrayFrekansSParam = new ArrayList();
            ArrayFrekansEE = new ArrayList();
            ArrayFrekansCF = new ArrayList();
            ArrayFrekansCIS = new ArrayList();
            ArrayFrekansNoise = new ArrayList();

            ArrayS11Reel = new ArrayList();
            ArrayS11ReelUnc = new ArrayList();
            ArrayS11Complex = new ArrayList();
            ArrayS11ComplexUnc = new ArrayList();
            ArrayS11Lin = new ArrayList();
            ArrayS11LinUnc = new ArrayList();
            ArrayS11LinPhase = new ArrayList();
            ArrayS11LinPhaseUnc = new ArrayList();
            ArrayS11Log = new ArrayList();
            ArrayS11LogUnc = new ArrayList();
            ArrayS11LogPhase = new ArrayList();
            ArrayS11LogPhaseUnc = new ArrayList();
            ArrayS11SWR = new ArrayList();
            ArrayS11SWRUnc = new ArrayList();


            ArrayS12Reel = new ArrayList();
            ArrayS12ReelUnc = new ArrayList();
            ArrayS12Complex = new ArrayList();
            ArrayS12ComplexUnc = new ArrayList();
            ArrayS12Lin = new ArrayList();
            ArrayS12LinUnc = new ArrayList();
            ArrayS12LinPhase = new ArrayList();
            ArrayS12LinPhaseUnc = new ArrayList();
            ArrayS12Log = new ArrayList();
            ArrayS12LogUnc = new ArrayList();
            ArrayS12LogPhase = new ArrayList();
            ArrayS12LogPhaseUnc = new ArrayList();


            ArrayS21Reel = new ArrayList();
            ArrayS21ReelUnc = new ArrayList();
            ArrayS21Complex = new ArrayList();
            ArrayS21ComplexUnc = new ArrayList();
            ArrayS21Lin = new ArrayList();
            ArrayS21LinUnc = new ArrayList();
            ArrayS21LinPhase = new ArrayList();
            ArrayS21LinPhaseUnc = new ArrayList();
            ArrayS21Log = new ArrayList();
            ArrayS21LogUnc = new ArrayList();
            ArrayS21LogPhase = new ArrayList();
            ArrayS21LogPhaseUnc = new ArrayList();


            ArrayS22Reel = new ArrayList();
            ArrayS22ReelUnc = new ArrayList();
            ArrayS22Complex = new ArrayList();
            ArrayS22ComplexUnc = new ArrayList();
            ArrayS22Lin = new ArrayList();
            ArrayS22LinUnc = new ArrayList();
            ArrayS22LinPhase = new ArrayList();
            ArrayS22LinPhaseUnc = new ArrayList();
            ArrayS22Log = new ArrayList();
            ArrayS22LogUnc = new ArrayList();
            ArrayS22LogPhase = new ArrayList();
            ArrayS22LogPhaseUnc = new ArrayList();
            ArrayS22SWR = new ArrayList();
            ArrayS22SWRUnc = new ArrayList();

            ArrayEffiencyEEEE = new ArrayList();
            ArrayEffiencyEEEEUnc = new ArrayList();
            ArrayEffiencyEE_S11Reel = new ArrayList();
            ArrayEffiencyEE_S11ReelUnc = new ArrayList();
            ArrayEffiencyEE_S11Imag = new ArrayList();
            ArrayEffiencyEE_S11ImagUnc = new ArrayList();
            ArrayEffiencyRHO_EERho = new ArrayList();
            ArrayEffiencyRHO_EERhoUnc = new ArrayList();
            ArrayEffiencyEE_CFEE_CF = new ArrayList();
            ArrayEffiencyEE_CFEE_CFUnc = new ArrayList();

            ArrayCF_Cal_Factor = new ArrayList();
            ArrayCF_Cal_Factor_Unc = new ArrayList();
            ArrayCF_Imaginer = new ArrayList();
            ArrayCF_Imaginer_Unc = new ArrayList();
            ArrayCF_Reel = new ArrayList();
            ArrayCF_Reel_Unc = new ArrayList();
            ArrayCF_ReflectionCof = new ArrayList();
            ArrayCF_ReflectionCof_Unc = new ArrayList();

            ArrayCIS_Z_Position = new ArrayList();
            ArrayCIS_Z_Position_Unc = new ArrayList();
            ArrayCIS_ICOD = new ArrayList();
            ArrayCIS_ICOD_Unc = new ArrayList();
            ArrayCIS_OCID = new ArrayList();
            ArrayCIS_OCID_Unc = new ArrayList();

            ArrayNoiseENR = new ArrayList();
            ArrayNoiseENRUnc = new ArrayList();
            ArrayNoiseDCONRCLin = new ArrayList();
            ArrayNoiseDCONUpLimit = new ArrayList();
            ArrayNoiseDCONRCLinUnc = new ArrayList();
            ArrayNoiseDCONRCPhase = new ArrayList();
            ArrayNoiseDCONRCPhaseUnc = new ArrayList();
            ArrayNoiseDCONControl = new ArrayList();
            ArrayNoiseDCOFFRCLin = new ArrayList();
            ArrayDCOFFUpLimit = new ArrayList();
            ArrayNoiseDCOFFRCLinUnc = new ArrayList();
            ArrayNoiseDCOFFRCPhase = new ArrayList();
            ArrayNoiseDCOFFRCPhaseUnc = new ArrayList();
            ArrayNoiseDCOFFControl = new ArrayList();

        }

        public void ClearData()
        {
            ArrayFrekansSParam.Clear();
            ArrayFrekansEE.Clear();
            ArrayFrekansCIS.Clear();
            ArrayFrekansCF.Clear();
            ArrayFrekansNoise.Clear();


            ArrayS11Reel.Clear();
            ArrayS11ReelUnc.Clear();
            ArrayS11Complex.Clear();
            ArrayS11ComplexUnc.Clear();
            ArrayS11Lin.Clear();
            ArrayS11LinUnc.Clear();
            ArrayS11LinPhase.Clear();
            ArrayS11LinPhaseUnc.Clear();
            ArrayS11Log.Clear();
            ArrayS11LogUnc.Clear();
            ArrayS11LogPhase.Clear();
            ArrayS11LogPhaseUnc.Clear();
            ArrayS11SWR.Clear();
            ArrayS11SWRUnc.Clear();
            ArrayS12Reel.Clear();
            ArrayS12ReelUnc.Clear();
            ArrayS12Complex.Clear();
            ArrayS12ComplexUnc.Clear();
            ArrayS12Lin.Clear();
            ArrayS12LinUnc.Clear();
            ArrayS12LinPhase.Clear();
            ArrayS12LinPhaseUnc.Clear();
            ArrayS12Log.Clear();
            ArrayS12LogUnc.Clear();
            ArrayS12LogPhase.Clear();
            ArrayS12LogPhaseUnc.Clear();
            ArrayS21Reel.Clear();
            ArrayS21ReelUnc.Clear();
            ArrayS21Complex.Clear();
            ArrayS21ComplexUnc.Clear();
            ArrayS21Lin.Clear();
            ArrayS21LinUnc.Clear();
            ArrayS21LinPhase.Clear();
            ArrayS21LinPhaseUnc.Clear();
            ArrayS21Log.Clear();
            ArrayS21LogUnc.Clear();
            ArrayS21LogPhase.Clear();
            ArrayS21LogPhaseUnc.Clear();
            ArrayS22Reel.Clear();
            ArrayS22ReelUnc.Clear();
            ArrayS22Complex.Clear();
            ArrayS22ComplexUnc.Clear();
            ArrayS22Lin.Clear();
            ArrayS22LinUnc.Clear();
            ArrayS22LinPhase.Clear();
            ArrayS22LinPhaseUnc.Clear();
            ArrayS22Log.Clear();
            ArrayS22LogUnc.Clear();
            ArrayS22LogPhase.Clear();
            ArrayS22LogPhaseUnc.Clear();
            ArrayS22SWR.Clear();
            ArrayS22SWRUnc.Clear();
            ArrayEffiencyEEEE.Clear();
            ArrayEffiencyEEEEUnc.Clear();
            ArrayEffiencyEE_S11Reel.Clear();
            ArrayEffiencyEE_S11ReelUnc.Clear();
            ArrayEffiencyEE_S11Imag.Clear();
            ArrayEffiencyEE_S11ImagUnc.Clear();
            ArrayEffiencyRHO_EERho.Clear();
            ArrayEffiencyRHO_EERhoUnc.Clear();
            ArrayEffiencyEE_CFEE_CF.Clear();
            ArrayEffiencyEE_CFEE_CFUnc.Clear();
            ArrayCF_Cal_Factor.Clear();
            ArrayCF_Cal_Factor_Unc.Clear();
            ArrayCF_Imaginer.Clear();
            ArrayCF_Imaginer_Unc.Clear();
            ArrayCF_Reel.Clear();
            ArrayCF_Reel_Unc.Clear();
            ArrayCF_ReflectionCof.Clear();
            ArrayCF_ReflectionCof_Unc.Clear();

            ArrayCIS_Z_Position.Clear();
            ArrayCIS_Z_Position_Unc.Clear();
            ArrayCIS_ICOD.Clear();
            ArrayCIS_ICOD_Unc.Clear();
            ArrayCIS_OCID.Clear();
            ArrayCIS_OCID_Unc.Clear();

            ArrayNoiseENR.Clear();
            ArrayNoiseENRUnc.Clear();
            ArrayNoiseDCONRCLin.Clear();
            ArrayNoiseDCONUpLimit.Clear();
            ArrayNoiseDCONRCLinUnc.Clear();
            ArrayNoiseDCONRCPhase.Clear();
            ArrayNoiseDCONRCPhaseUnc.Clear();
            ArrayNoiseDCONControl.Clear();
            ArrayNoiseDCOFFRCLin.Clear();
            ArrayDCOFFUpLimit.Clear();
            ArrayNoiseDCOFFRCLinUnc.Clear();
            ArrayNoiseDCOFFRCPhase.Clear();
            ArrayNoiseDCOFFRCPhaseUnc.Clear();
            ArrayNoiseDCOFFControl.Clear();
        }

    }
}
