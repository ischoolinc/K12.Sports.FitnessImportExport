using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Sports.FitnessImportExport.Log
{
    /// <summary>
    /// Log記錄轉換者
    /// </summary>
    public class LogTransfer
    {
        Dictionary<string, LogValue> _LogValueDict;

        public LogTransfer()
        {
            _LogValueDict = new Dictionary<string, LogValue>();
        }

        /// <summary>
        /// 設定Log值
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Value"></param>
        public void SetLogValue(string Name, string Value)
        {
            if (!_LogValueDict.ContainsKey(Name))
            {
                LogValue lv = new LogValue();
                lv.Name = Name;
                lv.OldValue = Value;
                lv.NewValue = "";
                _LogValueDict.Add(Name, lv);
            }
            else
            {
                _LogValueDict[Name].NewValue = Value;
            }
        }

        /// <summary>
        /// 取得LogValueList
        /// </summary>
        /// <returns></returns>
        public List<LogValue> getLogValueList()
        {
            return _LogValueDict.Values.ToList();
        }

        /// <summary>
        /// 請除LogValueList值
        /// </summary>
        public void Clear()
        {
            _LogValueDict.Clear();
        }

        /// <summary>
        /// 改變Log
        /// </summary>
        /// <param name="ActionBy"></param>
        /// <param name="Action"></param>
        /// <param name="BeforeString"></param>
        /// <param name="AfterString"></param>
        /// <param name="LogValueList"></param>
        /// <param name="targetCategory"></param>
        /// <param name="targetID"></param>
        public void SaveChangeLog(string ActionBy, string Action, string BeforeString, string AfterString, string targetCategory, string targetID)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(BeforeString);
            foreach (LogValue lv in GetDiffentLogValueList())
                sb.AppendLine(lv.getChangeString1());
            sb.Append(AfterString);
            FISCA.LogAgent.ApplicationLog.Log(ActionBy, Action, targetCategory, targetID, sb.ToString());
        }

        /// <summary>
        /// 新增 Log
        /// </summary>
        /// <param name="ActionBy"></param>
        /// <param name="Action"></param>
        /// <param name="BeforeString"></param>
        /// <param name="AfterString"></param>
        /// <param name="LogValueList"></param>
        /// <param name="targetCategory"></param>
        /// <param name="targetID"></param>
        public void SaveInsertLog(string ActionBy, string Action, string BeforeString, string AfterString, string targetCategory, string targetID)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(BeforeString);
            foreach (LogValue lv in GetDiffentLogValueList())
                sb.AppendLine(lv.getInsertString1());

            sb.Append(AfterString);
            FISCA.LogAgent.ApplicationLog.Log(ActionBy, Action, targetCategory, targetID, sb.ToString());
        }

        /// <summary>
        /// 刪除Log
        /// </summary>
        /// <param name="ActionBy"></param>
        /// <param name="Action"></param>
        /// <param name="BeforeString"></param>
        /// <param name="AfterString"></param>
        /// <param name="LogValueList"></param>
        /// <param name="targetCategory"></param>
        /// <param name="targetID"></param>
        public void SaveDeleteLog(string ActionBy, string Action, string BeforeString, string AfterString, string targetCategory, string targetID)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(BeforeString);
            foreach (LogValue lv in GetDiffentLogValueList())
                sb.AppendLine(lv.getDeleteString1());

            sb.Append(AfterString);
            FISCA.LogAgent.ApplicationLog.Log(ActionBy, Action, targetCategory, targetID, sb.ToString());
        }

        /// <summary>
        /// 單筆儲存
        /// </summary>
        /// <param name="ActionBy"></param>
        /// <param name="Action"></param>
        /// <param name="targetCategory"></param>
        /// <param name="targetID"></param>
        /// <param name="LogData"></param>
        public void SaveLog(string ActionBy, string Action, string targetCategory, string targetID, StringBuilder LogData)
        {
            FISCA.LogAgent.ApplicationLog.Log(ActionBy, Action, targetCategory, targetID, LogData.ToString());
        }

        /// <summary>
        /// 取得差異 log
        /// </summary>
        /// <returns></returns>
        public List<LogValue> GetDiffentLogValueList()
        {
            List<LogValue> retVal = new List<LogValue>();
            foreach (LogValue lv in getLogValueList())
            {
                if (lv.OldValue != lv.NewValue)
                    retVal.Add(lv);
            }
            return retVal;
        }
    }
}
