using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Globalization;

namespace SigmaSure
{
    public enum ResultType
    {
        None,
        ValueDouble,
        ValueString
    }

    public class UnitReport
    {
        public DateTime starttime;
        public DateTime endtime;
        public String mode;
        public String version;
        public _Station Station;
        public _Operator Operator;
        public _Cathegory Cathegory;
        public _TestRun TestRun;
        public _PropertyUR[] Properties;

        public UnitReport()
        {
            this.starttime = DateTime.Now;
            this.endtime = DateTime.Now;
            this.mode = "";
            this.version = "";
            this.Station = new _Station();
            this.Operator = new _Operator();
            this.Cathegory = new _Cathegory();
            this.TestRun = new _TestRun();
            this.Properties = new _PropertyUR[0];
            this.AddProperty("Memory", "N/A");
        }

        public UnitReport(DateTime starttime, DateTime endtime, String mode, String version)
        {
            this.starttime = starttime;
            this.endtime = endtime;
            this.mode = mode;
            this.version = version;
            this.Station = new _Station();
            this.Operator = new _Operator();
            this.Cathegory = new _Cathegory();
            this.TestRun = new _TestRun();
            this.Properties = new _PropertyUR[0];
            this.AddProperty("Memory", "N/A");
        }

        public UnitReport(DateTime starttime, DateTime endtime, String mode, String version, Boolean TestNumberPrefix)
        {
            this.starttime = starttime;
            this.endtime = endtime;
            this.mode = mode;
            this.version = version;
            this.Station = new _Station();
            this.Operator = new _Operator();
            this.Cathegory = new _Cathegory();
            this.TestRun = new _TestRun();
            this.Properties = new _PropertyUR[0];
            this.AddProperty("Memory", "N/A");
            this.TestNumberPrefix = TestNumberPrefix;
        }

        private Boolean TestNumberPrefix = true;

        public Array GetXMLReport()
        {
            return GenerateXML();
        }

        public Array GetXMLReport(String DirectoryToSave, Boolean Rewrite)
        {
            Array ar_buffer = GenerateXML();

            if (!Directory.Exists(DirectoryToSave))
            {
                Directory.CreateDirectory(DirectoryToSave);
            }
            String XMLfileName = String.Concat("SigmaProbe_", this.Cathegory.Product.SerialNo, "_", this.starttime.Hour.ToString("D2"), this.starttime.Minute.ToString("D2"), this.starttime.Second.ToString("D2"), "_", this.starttime.Month.ToString("D2"), this.starttime.Day.ToString("D2"), this.starttime.Year.ToString(), ".xml");
            if (File.Exists(String.Concat(DirectoryToSave, XMLfileName)))
            {
                if (Rewrite)
                {
                    File.Delete(String.Concat(DirectoryToSave, XMLfileName));
                    StreamWriter sw = new StreamWriter(String.Concat(DirectoryToSave, XMLfileName), true);
                    foreach (String actLine in ar_buffer)
                    {
                        sw.WriteLine(actLine);
                    }
                    sw.Close();
                }
                else
                {
                    throw new Exception("File exists.");
                }
            }  
            else
            {
                StreamWriter sw = new StreamWriter(String.Concat(DirectoryToSave, XMLfileName), true);
                foreach (String actLine in ar_buffer)
                {
                    sw.WriteLine(actLine);
                }
                sw.Close();
            } 
            return ar_buffer;
        }            

        public void AddProperty(String name, String value)
        {            
            foreach (_PropertyUR actProp in this.Properties)
            {
                if (actProp.name == name)
                {
                    actProp.value = value;
                    return;
                }
            }
            Array.Resize(ref this.Properties, this.Properties.Length + 1);
            this.Properties.SetValue(new _PropertyUR(name, value), this.Properties.Length - 1);
        }

        private String[] GenerateXML()
        {
            Boolean b_founded = false;
            foreach (_PropertyUR actProp in this.Properties)
            {
                if (actProp.name == "Work Order")
                {
                    if (this.Cathegory.Product.SerialNo.Length < 8)
                    {
                        this.Cathegory.Product.SerialNo = String.Concat(actProp.value, this.Cathegory.Product.SerialNo);
                    }
                    else
                    {
                        if (this.Cathegory.Product.SerialNo.Substring(0,8) != actProp.value
                            && this.Cathegory.Product.SerialNo.Substring(0, 7) != actProp.value.Substring(1,7))
                        {
                            this.Cathegory.Product.SerialNo = String.Concat(actProp.value, this.Cathegory.Product.SerialNo);
                        }
                        else
                        {

                            if ((this.Cathegory.Product.SerialNo.Length < 8) && (this.Cathegory.Product.SerialNo.Substring(0, 7) == actProp.value.Substring(1, 7)))
                            {
                                this.Cathegory.Product.SerialNo = String.Concat("0", this.Cathegory.Product.SerialNo);
                            }
                        }
                    }
                    b_founded = true;
                }
            }
            if (!b_founded)
            {
                throw new Exception("Work Order property is missing");                
            }

            String[] retArray = { "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>" };
            if (this.endtime == null) throw new Exception("Missing endtime.");
            if (this.starttime == null) throw new Exception("Missing starttime.");
            if (this.mode == null) throw new Exception("Missing mode.");
            retArray = this.AppendLine(retArray, String.Concat("<UnitReport end-time=\"", this.GetFormatedDatetime(this.endtime), "\" mode=\"", this.mode, "\" start-time=\"", this.GetFormatedDatetime(this.starttime), "\">"));
            if (this.Station.guid == null) throw new Exception("Missing station guid.");
            if (this.Station.name == null) throw new Exception("Missing station name.");
            retArray = this.AppendLine(retArray, String.Concat("\t<Station guid=\"", this.Station.guid, "\" name=\"", this.Station.name, "\"/>"));
            if (this.Operator.name == null) throw new Exception("Missing operator name.");
            retArray = this.AppendLine(retArray, String.Concat("\t<Operator name=\"", this.Operator.name, "\"/>"));
            if (this.Cathegory.name == null) throw new Exception("Missing cathegory name.");
            retArray = this.AppendLine(retArray, String.Concat("\t<Category name=\"", this.Cathegory.name, "\">"));
            if (this.Cathegory.Product.PartNo == null) throw new Exception("Missing product part number.");
            if (this.Cathegory.Product.SerialNo == null) throw new Exception("Missing product serial number.");
            retArray = this.AppendLine(retArray, String.Concat("\t\t<Product part-no=\"", this.Cathegory.Product.PartNo, "\" serial-no=\"", this.Cathegory.Product.SerialNo, "\"/>"));
            retArray = this.AppendLine(retArray, String.Concat("\t</Category>"));
            if (this.TestRun.name== null) throw new Exception("Missing testrun name.");
            if (this.TestRun.grade == null) throw new Exception("Missing testrun grade.");

            if (this.TestRun.grade == "")
            {
                this.TestRun.grade = "PASS";
                foreach (_TestRunStep actTRStep in this.TestRun.Testruns)
                {
                    if (actTRStep.grade != "PASS")
                    {
                        this.TestRun.grade = actTRStep.grade;
                        break;
                    }
                }
            }

            retArray = this.AppendLine(retArray, String.Concat("\t<TestRun end-time=\"", this.GetFormatedDatetime(this.endtime), "\" grade=\"", this.TestRun.grade, "\" name=\"", this.TestRun.name, "\" start-time=\"", this.GetFormatedDatetime(this.starttime), "\">"));

            Int16 n_counter = 0;

            foreach (_TestRunStep actTR in this.TestRun.Testruns)
            {
                n_counter++;
                if (this.TestNumberPrefix)
                {
                    actTR.name = String.Concat("Test ", n_counter.ToString(), ": ", actTR.Result.Property.name);
                }
                retArray = this.AppendLine(retArray, String.Concat("\t\t<TestRun end-time=\"", this.GetFormatedDatetime(actTR.endtime), "\" grade=\"", actTR.grade, "\" name=\"", actTR.name, "\" start-time=\"", this.GetFormatedDatetime(actTR.starttime), "\">"));
                retArray = this.AppendLine(retArray, String.Concat("\t\t\t<Result>"));
                retArray = this.AppendLine(retArray, String.Concat("\t\t\t\t<Property>"));

                String str_LSL = actTR.Result.Property.lsl.ToString().Replace(',', '.');
                while ((str_LSL.IndexOf('.') != -1) && (str_LSL.Substring(str_LSL.Length - 1) == "0")) str_LSL = str_LSL.Substring(0, str_LSL.Length - 1);

                String str_D_USL = actTR.Result.Property.d_usl.ToString().Replace(',', '.');
                while ((str_D_USL.IndexOf('.') != -1) && (str_D_USL.Substring(str_D_USL.Length - 1) == "0")) str_D_USL = str_D_USL.Substring(0, str_D_USL.Length - 1);

                String str_VALUE = actTR.Result.Property.measValue.ToString().Replace(',', '.');
                while ((str_VALUE.IndexOf('.') != -1) && (str_VALUE.Substring(str_VALUE.Length - 1) == "0")) str_VALUE = str_VALUE.Substring(0, str_VALUE.Length - 1);

                if (actTR.Result.Property.ResultType == ResultType.ValueDouble)
                {
                    if (actTR.Result.Property.d_usl == -9999.9999)
                    {
                        retArray = this.AppendLine(retArray, String.Concat("\t\t\t\t\t<ValueDouble lsl=\"", str_LSL, "\" name=\"", actTR.Result.Property.name, "\" uom=\"", actTR.Result.Property.uom, "\" usl=\"", "\">", str_VALUE, "</ValueDouble>"));
                    }
                    else
                    {
                        retArray = this.AppendLine(retArray, String.Concat("\t\t\t\t\t<ValueDouble lsl=\"", str_LSL, "\" name=\"", actTR.Result.Property.name, "\" uom=\"", actTR.Result.Property.uom, "\" usl=\"", str_D_USL, "\">", str_VALUE, "</ValueDouble>"));
                    }
                }
                else if (actTR.Result.Property.ResultType == ResultType.ValueString)
                {
                    retArray = this.AppendLine(retArray, String.Concat("\t\t\t\t\t<ValueString name=\"", actTR.Result.Property.name, "\" uom=\"", actTR.Result.Property.uom, "\" usl=\"", actTR.Result.Property.str_usl, "\">", actTR.Result.Property.measString, "</ValueString>"));
                }
                retArray = this.AppendLine(retArray, String.Concat("\t\t\t\t</Property>"));
                retArray = this.AppendLine(retArray, String.Concat("\t\t\t</Result>"));
                retArray = this.AppendLine(retArray, String.Concat("\t\t</TestRun>"));
            }
            retArray = this.AppendLine(retArray, String.Concat("\t</TestRun>"));

            retArray = this.AppendLine(retArray, String.Concat("\t<Property>"));
            foreach (_PropertyUR actProperty in this.Properties)
            {
                retArray = this.AppendLine(retArray, String.Concat("\t\t<ValueString name=\"", actProperty.name, "\">", actProperty.value, "</ValueString>"));
            }
            retArray = this.AppendLine(retArray, String.Concat("\t</Property>"));

            retArray = this.AppendLine(retArray, String.Concat("</UnitReport>"));



            return retArray;
        }
    

        private String[] AppendLine(String[] Lines, String LineToAppend)
        {
            Array.Resize(ref Lines, Lines.Length + 1);
            Lines.SetValue(LineToAppend, Lines.Length - 1);
            return Lines;
        }

        private String GetFormatedDatetime(DateTime dt)
        {
            String retValue = String.Concat(dt.Year, "-", dt.Month.ToString("D2"), "-", dt.Day.ToString("D2"), "T",
                dt.Hour.ToString("D2"), ":", dt.Minute.ToString("D2"), ":", dt.Second.ToString("D2"));
            String curTZ = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).ToString();
            if (curTZ.Substring(0,1) == "-")
            {
                retValue = String.Concat(retValue, "-");
                curTZ = curTZ.Substring(1);
            }
            else
            {
                retValue = String.Concat(retValue, "+");
            }
            retValue = string.Concat(retValue, curTZ.Substring(0, 5));
            return retValue;           
        }
    }

    public class _Station
    {
        public String guid;
        public String name;

        public _Station()
        {
            this.guid = "";
            this.name = "";
        }

        public _Station(String guid, String name)
        {
            this.guid = guid;
            this.name = name;
        }
    }

    public class _Operator
    {
        public String name = "";

        public _Operator()
        {
            this.name = "";
        }

        public _Operator(String name)
        {
            this.name = name;
        }
    }

    public class _Cathegory
    {
        public String name;
        public _Product Product;        

        public _Cathegory()
        {
            this.name = "";
            this.Product = new _Product();
        }

        public _Cathegory(String name)
        {
            this.name = name;
            this.Product = new _Product();
        }
    }

    public class _Product
    {
        public String PartNo;
        public String SerialNo;
        
        public _Product()
        {
            this.PartNo = "";
            this.SerialNo = "";
        }

        public _Product(String PartNo, String SerialNo)
        {
            this.PartNo = PartNo;

            this.SerialNo = SerialNo;
        }
    }

    public class _PropertyUR
    {
        public String name;
        public String value;

        public _PropertyUR()
        {
            this.name = "";
            this.value = "";
        }

        public _PropertyUR(String name, String value)
        {
            this.name = name;
            this.value = value;
            if (this.name == "Work Order")
            {
                while (this.value.Length < 8)
                {
                    this.value = String.Concat("0", this.value);
                }
            }
        }
    }

    public class _TestRun
    {
        public String grade;
        public DateTime starttime;
        public DateTime endtime;
        public String name;
        public _TestRunStep[] Testruns;        

        public _TestRun()
        {
            this.grade = "";
            this.starttime = DateTime.Now;
            this.endtime = DateTime.Now;
            this.name = "";
            this.Testruns = new _TestRunStep[0];
        }

        public _TestRun(String grade, DateTime starttime, DateTime endtime, String name)
        {
            this.grade = grade;
            this.starttime = starttime;
            this.endtime = endtime;
            this.name = name;
        }

        public void AddTestRunChild(_TestRunStep TestRunToAdd)
        {
            Array.Resize(ref this.Testruns, this.Testruns.Length + 1);
            this.Testruns.SetValue(TestRunToAdd, this.Testruns.Length - 1);
        }

        public void AddTestRunChild(String name, DateTime starttime, DateTime endtime, String grade, String uom, Double measValue, Double lsl, Double usl)
        {
            _TestRunStep TRStepToAdd = new _TestRunStep();
                        
            TRStepToAdd.name = name;
            TRStepToAdd.starttime = starttime;
            TRStepToAdd.endtime = endtime;
            if (grade != "") TRStepToAdd.grade = grade;
            else
            {
                if ((lsl > measValue) || (usl < measValue)) TRStepToAdd.grade = "FAIL";
                else TRStepToAdd.grade = "PASS";
                
            }
            TRStepToAdd.Result.Property.ResultType = ResultType.ValueDouble;
            TRStepToAdd.Result.Property.name = name;
            TRStepToAdd.Result.Property.uom = uom;
            TRStepToAdd.Result.Property.measValue = measValue;
            TRStepToAdd.Result.Property.lsl = lsl;
            TRStepToAdd.Result.Property.d_usl = usl;
            
            this.AddTestRunChild(TRStepToAdd);
        }

        public void AddTestRunChild(String name, DateTime starttime, DateTime endtime, String grade, String uom, Double measValue, Double lsl)
        {
            _TestRunStep TRStepToAdd = new _TestRunStep();

            TRStepToAdd.name = name;
            TRStepToAdd.starttime = starttime;
            TRStepToAdd.endtime = endtime;
            TRStepToAdd.grade = grade;
            TRStepToAdd.Result.Property.ResultType = ResultType.ValueDouble;
            TRStepToAdd.Result.Property.name = name;
            TRStepToAdd.Result.Property.uom = uom;
            TRStepToAdd.Result.Property.measValue = measValue;
            TRStepToAdd.Result.Property.lsl = lsl;
            TRStepToAdd.Result.Property.d_usl = -9999.9999;
            this.AddTestRunChild(TRStepToAdd);
        }

        public void AddTestRunChild(String name, DateTime starttime, DateTime endtime, String grade, String measString, String uom, String usl)
        {
            _TestRunStep TRStepToAdd = new _TestRunStep();

            TRStepToAdd.name = name;
            TRStepToAdd.starttime = starttime;
            TRStepToAdd.endtime = endtime;
            TRStepToAdd.grade = grade;
            TRStepToAdd.Result.Property.ResultType = ResultType.ValueString;
            TRStepToAdd.Result.Property.name = name;
            TRStepToAdd.Result.Property.measString = measString;
            TRStepToAdd.Result.Property.uom = uom;
            TRStepToAdd.Result.Property.str_usl = usl;
            this.AddTestRunChild(TRStepToAdd);
        }

    }

    public class _TestRunStep
    {
        public DateTime endtime;
        public DateTime starttime;
        public String name;
        public String grade;
        public _Result Result;
        
        public _TestRunStep()
        {
            this.endtime = DateTime.Now;
            this.starttime = DateTime.Now;
            this.name = "";
            this.grade = "";
            this.Result = new _Result();
        }

        public _TestRunStep(DateTime starttime, DateTime endtime, String name, String grade)
        {
            this.starttime = starttime;
            this.endtime = endtime;
            this.name = name;
            this.grade = grade;
            this.Result = new _Result();
        }
    }

    public class _Result
    {
        public _PropertyTR Property;

        public _Result()
        {
            this.Property = new _PropertyTR();
        }

        public _Result(_PropertyTR Property)
        {
            this.Property = Property;
        }
    }

    public class _PropertyTR
    {
        public ResultType ResultType = ResultType.None;

        public String name;
        public Double measValue;
        public Double lsl;
        public Double d_usl;

        public String measString;
        public String uom;
        public String str_usl;

        public _PropertyTR()
        {
            this.ResultType = ResultType.None;
            this.name = "";
            this.measValue = 0.0;
            this.lsl = 0.0;
            this.d_usl = 0.0;
            this.measString = "";
            this.uom = "";
            this.str_usl = "";
        }

        public _PropertyTR(String name, Double measValue, Double lsl, Double usl)
        {
            this.ResultType = ResultType.ValueDouble;
            this.name = name;
            this.measValue = measValue;
            this.lsl = lsl;
            this.d_usl = usl;
            this.measString = "";
            this.uom = "";
            this.str_usl = "";
        }

        public _PropertyTR(String name, String measString, String uom, String usl)
        {
            this.ResultType = ResultType.ValueString;
            this.name = name;
            this.measString = measString;
            this.uom = uom;
            this.str_usl = usl;
            this.measValue = 0.0;
            this.lsl = 0.0;
            this.d_usl = 0.0;            
        }
    }    
}
