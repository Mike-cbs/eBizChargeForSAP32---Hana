using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;


public class CCCardAcct
{
    public int DocEntry { get; set; }
    public string U_GroupName { get; set; }
    public string U_CardType { get; set; }
    public string U_CardName { get; set; }
    public string U_AcctCode { get; set; }
    public string U_CardCode { get; set; }
    public string U_SourceKey { get; set; }
    public string U_Currency { get; set; }
    public string U_Pin { get; set; }
    public string U_BranchID { get; set; }
}
public partial class SAP
{

    public List<CCCardAcct> cardAcctList = new List<CCCardAcct>();
    /*
    const string CC_USD = "840";
    const string CC_CAN = "124";
    const string CC_EUR = "978";
    public string getCurrencyCode(string currency)
    {
        if (cfgIsSandbox == "Y")
            return CC_USD;

        switch (currency)
        {

            case "USD":
                return CC_USD;
            case "CAN":
                return CC_CAN;
            case "$":
                return CC_USD;
            case "CAD":
                return CC_CAN;
            case "EUR":
                return CC_EUR;
        }
        return CC_USD;
    }
     * */
    public void createAcctList()
    {


        string sql = "select \"DocEntry\",\"U_GroupName\",\"U_CardType\",\"U_CardName\",\"U_AcctCode\",\"U_CardCode\",\"U_SourceKey\",\"U_Pin\", \"U_Currency\", \"U_BranchID\" from \"@CCCARDACCT\"";
        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        try
        {
            initSetupCardAcct();
            initUpdateCardAcctWithOCRC("v", "isa");
            initUpdateCardAcctWithOCRC("m", "aster");
            initUpdateCardAcctWithOCRC("a", "mer");
            initUpdateCardAcctWithOCRC("ds", "iscover");
            initUpdateCardAcctWithOCRC("eCheck", "eCheck");
            string t = "Card name added to List: ";
            cardAcctList = new List<CCCardAcct>();
            oRS.DoQuery(sql);

            while (!oRS.EoF)
            {
                string g = (string)oRS.Fields.Item(1).Value;
                if (g.ToUpper().IndexOf("NOT USE") == -1)
                {

                    CCCardAcct acct = new CCCardAcct();
                    acct.DocEntry = (int)oRS.Fields.Item(0).Value;
                    acct.U_GroupName = (string)oRS.Fields.Item(1).Value;
                    acct.U_CardType = (string)oRS.Fields.Item(2).Value;
                    acct.U_CardName = (string)oRS.Fields.Item(3).Value;
                    acct.U_AcctCode = (string)oRS.Fields.Item(4).Value;
                    acct.U_CardCode = (string)oRS.Fields.Item(5).Value;
                    acct.U_SourceKey = (string)oRS.Fields.Item(6).Value;
                    acct.U_Pin = (string)oRS.Fields.Item(7).Value;
                    acct.U_Currency = (string)oRS.Fields.Item(8).Value;
                    acct.U_BranchID = (string)oRS.Fields.Item(9).Value;
                    cardAcctList.Add(acct);
                    t = t + acct.U_CardName + "," + acct.U_GroupName + "\r\n";
                }
                oRS.MoveNext();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
            errorLog(sql);
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            oRS = null;
        }
    }
    public string getCardName()
    {
        string str = "";
        try
        {
            var q = from x in cardAcctList
                    orderby x.U_CardName
                    select x;
            foreach (CCCardAcct acct in q)
            {
                str = acct.U_CardName + "," + str;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        if (str.Length > 0)
            str = str.Substring(0, str.Length - 1);
        return str;
    }
    public string getCardCode(string cardname)
    {


        try
        {
            var q = from x in cardAcctList
                    where x.U_CardName == cardname
                    select x;
            foreach (CCCardAcct acct in q)
            {
                return acct.U_CardCode;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return getCardCode();
    }

    public string getCardCode()
    {

        try
        {
            var q = from x in cardAcctList
                    select x;
            foreach (CCCardAcct acct in q)
            {
                return acct.U_CardCode;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return "0";
    }

    public string getCardName(string grp, string cardtype, string currency, ref string cardcode, string branchid = "")
    {
        if (branchid == "")
            branchid = getBranchID();
        if (cardAcctList.Count == 0)
        {
            createAcctList();
        }
        string str = "";

        try
        {
            string traceStr = string.Format("getCardName: grp={0}, type={1}, currency={2}; branchid={3}; ", grp, cardtype, currency, branchid);
            if (branchid != "0")
            {
                var q1 = from x in cardAcctList
                         where x.U_BranchID == branchid
                         select x;
                foreach (CCCardAcct acct in q1)
                {
                    trace(traceStr + string.Format("Found Card using branch ID: branchid = {0}, sourceKey={1}", acct.U_BranchID, acct.U_SourceKey));
                    cardcode = acct.U_CardCode;
                    return acct.U_CardName;
                }
            }
            if (grp != "" && grp != null)
            {
                var q1 = from x in cardAcctList
                         where x.U_GroupName == grp && x.U_CardType == "" && x.U_GroupName != ""
                         select x;
                foreach (CCCardAcct acct in q1)
                {
                    trace(traceStr + string.Format("Found Card using group: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                    cardcode = acct.U_CardCode;
                    return acct.U_CardName;
                }
            }
            var q2 = from x in cardAcctList
                     where (x.U_CardType.ToLower() == cardtype.ToLower() || x.U_CardType == "") && x.U_Currency == currency && x.U_Currency != ""
                     select x;
            foreach (CCCardAcct acct in q2)
            {
                trace(traceStr + string.Format("Found Card using currency: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                cardcode = acct.U_CardCode;
                return acct.U_CardName;
            }
            var q3 = from x in cardAcctList
                     where (x.U_CardType.ToLower() == cardtype.ToLower() || x.U_CardType == "") && (x.U_GroupName == null || x.U_GroupName == "") && x.U_Currency == ""
                     select x;
            foreach (CCCardAcct acct in q3)
            {
                trace(traceStr + string.Format("Found Card using card type: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                cardcode = acct.U_CardCode;
                return acct.U_CardName;
            }
            var q4 = from x in cardAcctList
                     select x;
            foreach (CCCardAcct acct in q4)
            {
                trace(traceStr + string.Format("Found Card no criteria: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                cardcode = acct.U_CardCode;
                return acct.U_CardName;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        str = getCardName();
        return str;
    }


    public void getPinSKFromCardAcct(string cardCode, ref string pin, ref string sourceKey)
    {

        try
        {
            try
            {
                var q = from x in cardAcctList
                        where x.U_CardCode == cardCode
                        select x;
                foreach (CCCardAcct acct in q)
                {
                    pin = acct.U_Pin;
                    sourceKey = acct.U_SourceKey;
                }
            }
            catch (Exception ex)
            {
                errorLog(ex);
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
    }
    public string getCurrencyCodeFormCard(string cardcode)
    {
        try
        {
            var q = from x in cardAcctList
                    where x.U_CardCode.ToString() == cardcode
                    select x;
            foreach (CCCardAcct acct in q)
            {
                return acct.U_Currency;
            }

        }
        catch (Exception)
        {

        }
        return "";
    }
    public string getCurrency(SAPbouiCOM.Form form)
    {
        if (form == null)
            return "";
        try
        {
            return getFormItemVal(form, cbCurrencyCode);


        }
        catch (Exception ex)
        {
            errorLog(ex);

        }
        return "";
    }
    public string getAccountName(string cardcode)
    {
        //if (cfgAcctNameField != "")
        //    return getSQLString(string.Format("select {0} from OCRD where CardCode = '{1}'", cfgAcctNameField, cardcode));
        
        return FindCustomerByID(cardcode);
    }
    public void getAllSecurityID(List<string> list)
    {
        try
        {
            var q = from x in cardAcctList
                    where x.U_SourceKey != ""
                    select x;
            foreach (CCCardAcct acct in q)
            {
                list.Add(acct.U_SourceKey);
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
    }
    public string getCardName(string grp, string cardtype, string currency, ref string cardcode, string AccountName, string branchid = "")
    {
        ;
        if (branchid == "")
            branchid = getBranchID();
        if (currency != null)
            currency = currency.Trim();
        if (cardAcctList.Count == 0)
        {
            createAcctList();
        }
        string str = "";

        string traceStr = string.Format("getCardName: branchid={3}, grp={0}, type={1}, currency={2}; account name= {4}; ", grp, cardtype, currency, branchid, AccountName);
        try
        {
            if (branchid != "0")
            {
                var q1 = from x in cardAcctList
                         where x.U_BranchID == branchid && x.U_BranchID != ""
                         select x;
                foreach (CCCardAcct acct in q1)
                {
                    cardcode = acct.U_AcctCode.ToString();

                    if (cardtype == "")
                    {
                        trace(traceStr + string.Format("Found Card using branch ID: branchid = {0}, sourceKey={1}", acct.U_BranchID, acct.U_SourceKey));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType.ToLower() == cardtype.ToLower())
                    {
                        trace(traceStr + string.Format("Found Card using branch ID: branchid = {0}, sourceKey={1}", acct.U_BranchID, acct.U_SourceKey));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType == "")
                    {
                        trace(traceStr + string.Format("Found Card using branch ID: branchid = {0}, sourceKey={1}", acct.U_BranchID, acct.U_SourceKey));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                }
                if (q1.Count() > 0)
                {
                    trace(traceStr + string.Format("Found Card using branch ID: branchid = {0}", branchid));
                    return q1.FirstOrDefault().U_CardName;
                }
            }
            if (grp != "" && grp != null)
            {
                var q1 = from x in cardAcctList
                         where x.U_GroupName.ToLower() == grp.ToLower() && x.U_GroupName != ""
                         select x;
                foreach (CCCardAcct acct in q1)
                {
                    cardcode = acct.U_AcctCode.ToString();
                    if (cardtype == "")
                    {
                        trace(traceStr + string.Format("Found Card using group: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType.ToLower() == cardtype.ToLower())
                    {
                        trace(traceStr + string.Format("Found Card using group: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType == "")
                    {
                        trace(traceStr + string.Format("Found Card using group: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }

                }
                if (q1.Count() > 0)
                {
                    trace(traceStr + string.Format("Found Card using group: grp={0}", grp));
                    return q1.FirstOrDefault().U_CardName;
                }
            }
            //AccountName
            if (AccountName != "" && AccountName != null)
            {
                var q1 = from x in cardAcctList
                         where x.U_Pin == AccountName && x.U_Pin != ""
                         select x;
                foreach (CCCardAcct acct in q1)
                {
                    cardcode = acct.U_AcctCode.ToString();
                    if (cardtype == "")
                    {
                        trace(traceStr + string.Format("Found Card using AccountName: AccountName={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", AccountName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType.ToLower() == cardtype.ToLower())
                    {
                        trace(traceStr + string.Format("Found Card using AccountName: AccountName={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", AccountName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType == "")
                    {
                        trace(traceStr + string.Format("Found Card using AccountName: AccountName={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", AccountName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));

                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }

                }
                if (q1.Count() > 0)
                {
                    trace(traceStr + string.Format("Found Card using AccountName: AccountName={0}", AccountName));

                    return q1.FirstOrDefault().U_CardName;
                }
            }
            if (currency != "" && currency != null)
            {
                var q2 = from x in cardAcctList
                         where x.U_Currency == currency && x.U_Currency != ""
                         select x;
                foreach (CCCardAcct acct in q2)
                {
                    trace(traceStr + string.Format("Found Card using currency: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                    if (cardtype == "")
                    {
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType.ToLower() == cardtype.ToLower())
                    {
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    else if (acct.U_CardType == "")
                    {
                        cardcode = acct.U_AcctCode.ToString();
                        return acct.U_CardName;
                    }
                    if (q2.Count() > 0)
                        return q2.FirstOrDefault().U_CardName;
                }
            }
            var q3 = from x in cardAcctList
                     where (x.U_CardType.ToLower() == cardtype.ToLower() || x.U_CardType == "") && (x.U_GroupName == null || x.U_GroupName == "")
                     select x;
            foreach (CCCardAcct acct in q3)
            {
                trace(traceStr + string.Format("Found Card using card type: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                cardcode = acct.U_AcctCode.ToString();
                return acct.U_CardName;
            }
            var q4 = from x in cardAcctList
                     select x;
            foreach (CCCardAcct acct in q4)
            {
                trace(traceStr + string.Format("Found Card no criteria: grp={0}, type={1}, currency={2}, card name:{3}, card Code:{4}", acct.U_GroupName, acct.U_CardType, acct.U_Currency, acct.U_CardName, acct.U_CardCode));
                cardcode = acct.U_AcctCode.ToString();
                return acct.U_CardName;
            }
        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        str = getCardName();
        return str;
    }

    public string geteConnectCardCode(string cardname)
    {


        try
        {
            trace("eConnect payment method: " + cardname);
            var q = from x in cardAcctList
                    where x.U_CardName.ToUpper() == ("eConnect-" + cardname).ToUpper()
                    select x;
            foreach (CCCardAcct acct in q)
            {
                return acct.U_CardCode.ToString();
            }
            var q2 = from x in cardAcctList
                     where x.U_CardName.ToUpper() == ("eConnect").ToUpper()
                     select x;
            foreach (CCCardAcct acct in q2)
            {
                return acct.U_CardCode.ToString();
            }

        }
        catch (Exception ex)
        {
            errorLog(ex);
        }
        return getCardCode(cardname);
    }
}

