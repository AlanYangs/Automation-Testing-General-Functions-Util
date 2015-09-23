//-------------------------------------------------------------------------------------
//Function Name : killProcess
//Author        : Alan.Yang
//Create Date   : May 11, 2015
//Last Modify   : 
//Description   : according to the name kill it's process
//Parameter     : [IN]ProcessName -- the name of process
//Return        : null
//-------------------------------------------------------------------------------------
function killProcess(ProcessName){

    //if(Sys.Process(ProcessName).Exists){
        var wshShell = new ActiveXObject("WScript.Shell");
        wshShell.Run("%comspec% /c taskkill /F /IM " + ProcessName +"*",true);
        wshShell = null;
        Sys.Refresh();
        Log.Message("The process [ "+ ProcessName + " ] has been killed.");
    //}
}

//-------------------------------------------------------------------------------------
//Function Name : exportReport
//Author        : Alan.Yang
//Create Date   : July 13, 2015
//Last Modify   : 
//Description   : export log report to specified path
//Parameter     : null
//Return        : Boolean
//-------------------------------------------------------------------------------------
function exportReport()
{
    var d = new Date();
    var strYear = d.getFullYear().toString();
    var strMonth = (d.getMonth()+1)<10 ? ".0"+(d.getMonth()+1) : "."+(d.getMonth()+1);
    var strDay = d.getDate()<10 ? ".0"+d.getDate() : "."+d.getDate();
    var strHour = d.getHours()<10 ? "-0"+d.getHours() : "-"+d.getHours();
    var strMinute = d.getMinutes()<10 ? "0"+d.getMinutes() : d.getMinutes(); 
    var strMainFolder = strYear + strMonth + strDay + strHour + strMinute;
    if(aqFile.Exists(gReportPath)){
        if(Log.SaveResultsAs(gReportPath + strMainFolder  + "\\TCLogs", 1)){
            Log.Message("Export execution Logs to path:" + gReportPath + strMainFolder);
            var HtmFiles = aqFileSystem.FindFiles(gReportPath + strMainFolder  + "\\TCLogs","*.htm");
            var strSumHtm = "";
            if(HtmFiles != null){
                while(HtmFiles.HasNext()){
                    var file = HtmFiles.Next();
                    Log.Message(file.Name);
                    if((file.Name).indexOf("index") == -1){
                        strSumHtm = file.Name;
                        break;
                    }
                }
            }
            if(strSumHtm != ""){
                var strHtml = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>Automated Testing Report</title></head><div><h1>SQL Navigator-BuildTest</h1></div><div id=\"auto\"><iframe src=\"" +gReportPath + strMainFolder + "\\TCLogs\\"+ strSumHtm +"\" height=\"100%\" width=\"100%\" frameborder=\"0\" scrolling =\"no\"><\\iframe><iframe src=\"" +gReportPath + strMainFolder + "\\TCLogs\\index.htm\" height=\"100%\" width=\"100%\" frameborder=\"0\" scrolling =\"no\"> <a href=\""+gReportPath + strMainFolder + "\\TCLogs\\index.htm\">您的浏览器版本太低，请点击这里访问页面内容</a> </iframe> </div></body></html>";
            }else{
                var strHtml = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>Automated Testing Report</title></head><div><h1>SQL Navigator-BuildTest</h1></div><div id=\"auto\"><iframe src=\"" +gReportPath + strMainFolder + "\\TCLogs\\index.htm\" height=\"100%\" width=\"100%\" frameborder=\"0\" scrolling =\"no\"> <a href=\""+gReportPath + strMainFolder + "\\TCLogs\\index.htm\">您的浏览器版本太低，请点击这里访问页面内容</a> </iframe> </div></body></html>";
            }
            if(aqFile.Create(gReportPath + strMainFolder + "\\Main.html") ==0){
                aqFile.WriteToTextFile(gReportPath + strMainFolder + "\\Main.html",strHtml,22,true);
            }
            
            return gReportPath + strMainFolder;
        }
        else{
            Log.Warning("Fail to export Logs to specified path.");
            return "";
        }
    }
    else{
        Log.Warning("Path ["+gReportPath+"] is not exists.");
        return "";
    }
}

function sendReport(strReportPath){
    if(strReportPath != ""){
        var strFromAddress = "XXX@163.com";
        var strFromHost = "smtp.163.com";
        var strFromName = "Automation Testing Team";
        var strToAddress = "Michael@auto.com;Vincent@autol.com;Alan@auto.com";
        var strSubject = "Automated Testing Report"
        var strBody = "<body><div><h1>SQLNavigator-BuildTest</h1></div><div>Build Name:"+ gStrBuildFile +"</div><br><div id=\"auto\"><iframe src=\"" +strReportPath+ "\\Main.html\" height=\"100%\" width=\"100%\" frameborder=\"0\" scrolling =\"no\"> <a href=\""+strReportPath+ "\\Main.html\">please click here to view the page content</a> </iframe> </div></body>"
        var strAttach = strReportPath+ "\\Main.html";
        try
        {
            SendEmail(strFromAddress,strToAddress,strSubject,strBody,strAttach);
        }
        catch (e)
        {
            Log.Warning(e.message);
        }
    }else{
        Log.Warning("Fail to send report by e-mail.");
    }
}

function SendEmail(mFrom, mTo, mSubject, mBody, mAttach){
    var i, schema, mConfig, mMessage;

    try
    {
      schema = "http://schemas.microsoft.com/cdo/configuration/";
      mConfig = Sys.OleObject("CDO.Configuration");
      mConfig.Fields.Item(schema + "sendusing") = 2; // cdoSendUsingPort
      mConfig.Fields.Item(schema + "smtpserver") = "smtp.163.com"; // SMTP server
      mConfig.Fields.Item(schema + "smtpserverport") = 25; // Port number
      mConfig.Fields.Item(schema + "smtpauthenticate") = 1; // Authentication mechanism
      mConfig.Fields.Item(schema + "sendusername") = "alan@163.com"; // User name (if needed)
      mConfig.Fields.Item(schema + "sendpassword") = "your password"; // User password (if needed)
      mConfig.Fields.Update();

      mMessage = Sys.OleObject("CDO.Message");
      mMessage.Configuration = mConfig;
      mMessage.From = mFrom;
      mMessage.To = mTo;
      mMessage.Subject = mSubject;
      mMessage.HTMLBody = mBody;

      aqString.ListSeparator = ",";
      for(i = 0; i < aqString.GetListLength(mAttach); i++)
          mMessage.AddAttachment(aqString.GetListItem(mAttach, i));
      mMessage.Send();
    }
    catch (exception)
    {
      Log.Error("E-mail cannot be sent.", exception.description);
      return false;
    }
    Log.Message("Message to <" + mTo + "> was successfully sent");
    return true;
}

//-------------------------------------------------------------------------------------
//Function Name : bObjExists
//Author        : Micahel.luo
//Create Date   : Mar 18, 2014
//Last Modify   : Alan.Yang  , May 11, 2015
//Description   : get the ini file Option value
//Parameter     : [IN]NameMappingItem -- Name Mapping object for Aliases name which you want to judge object exists
//Parameter     : [Optional IN] -- optional parameter nWaitTime s, by default is 30s
//Return        : boolean of object exists
//-------------------------------------------------------------------------------------
function bObjExists(NameMappingItem,intSecond){
    var waitSeconds = arguments[1] == undefined ? 30:intSecond;
    var counter = 1;
    while(!NameMappingItem.Exists)
    {
        Delay(1000);
        // Refresh the mapping information to see if the object has been recreated
        NameMappingItem.RefreshMappingInfo();
        if(counter > waitSeconds){
            //Log.Error("The object "+NameMappingItem +" is not exist, wait timeout.",null,pmNormal,null,Sys.Desktop);
            break;
        }
        counter++;
    }
    return NameMappingItem.Exists;
}

//-------------------------------------------------------------------------------------
//Function Name : execShortcuts
//Author        : Alan.Yang
//Create Date   : May 14, 2015
//Last Modify   : 
//Description   : executing the object's shortcuts
//Parameter     : [IN]objNameMapped -- the namemapped object
//Parameter     : [IN]strShortCut -- the keys of shortcut
//Return        : null
//-------------------------------------------------------------------------------------
function execShortcuts(objNameMapped,strShortcut){

    if(objNameMapped.Exists){
        objNameMapped.SetFocus();
        //spliting the string of shortcuts 
        for(i=0;i<strShortcut.length;i++)
        {
            objNameMapped.Keys(strShortcut.charAt(i));
            Log.Message("Typing shortcut:" + strShortcut.charAt(i));
            Delay(100);
        }
    } 
}

//-------------------------------------------------------------------------------------
//Function Name : trim
//Author        : Alan.Yang
//Create Date   : May 21, 2015
//Last Modify   : 
//Description   : clear the string's space
//Parameter     : [IN]s -- the source string
//Return        : String
//-------------------------------------------------------------------------------------
//clear the right and left space
function trim(s){
    return trimRight(trimLeft(s)); 
} 
//clear the left space
function trimLeft(s){ 
    if(s == null) { 
      return ""; 
    } 
    //var whitespace = new String(" \n\r"); 
    //var str = new String(s); 
    var whitespace = " \n\r";
    var str = s;
    if (whitespace.indexOf(str.charAt(0)) != -1) { 
      var j=0, i = str.length; 
      while (j < i && whitespace.indexOf(str.charAt(j)) != -1){ 
          j++; 
      } 
      str = str.substring(j, i); 
    } 
    return str; 
} 
//clear the right space 
function trimRight(s){ 
    if(s == null) return ""; 
    //var whitespace = new String(" \n\r"); 
    //var str = new String(s); 
    var whitespace = " \n\r";
    var str = s;
    if (whitespace.indexOf(str.charAt(str.length-1)) != -1){ 
      var i = str.length - 1; 
      while (i >= 0 && whitespace.indexOf(str.charAt(i)) != -1){ 
          i--; 
      } 
      str = str.substring(0, i+1); 
    } 
    return str; 
} 

//-------------------------------------------------------------------------------------
//Function Name : getNumString
//Author        : Alan.Yang
//Create Date   : May 29, 2015
//Last Modify   : 
//Description   : generating the number of specified strings
//Parameter     : [IN]strKey -- the specified string
//Parameter     : [IN]number -- need to generate string's number
//Return        : String
//-------------------------------------------------------------------------------------
function getNumString(strKey,number){
    if(number <= 0) return ""; 
    var Keys = "";
    for(i=0; i<number; i++){
        Keys += strKey;
    }
    return Keys;
}

//-------------------------------------------------------------------------------------
//Function Name : gotoAndExpandTree
//Author        : Alan.Yang
//Create Date   : June 9, 2015
//Last Modify   : 
//Description   : goto and expand the specified Tree path
//Parameter     : [IN]objTree -- the object Tree
//Parameter     : [IN]strPaths -- the path of tree，split by "-",such as 6-1-1
//Return        : Object, return last node in the specified tree path
//-------------------------------------------------------------------------------------
function gotoAndExpandTree(objTree, strPaths){
    
    strPaths = (arguments[1] == undefined || strPaths=="") ? 0 : strPaths;//deal with null and ""
    var objNode = objTree.wItems.Item(0);//root node
    var arrLevels = new Array();
    if(strPaths.indexOf("-")!= -1){
        arrLevels = strPaths.split("-");
        for(i=0; i<arrLevels.length; i++){
            arrLevels[i] = parseInt(arrLevels[i]);//transfer to integer
        }
    }
    else{
        arrLevels[0] = parseInt(strPaths);
    }
    var counter = 0;
    for(i=0; i<arrLevels.length; i++){
        var intPos = objTree.VScroll.Pos;
        objNode = (i == arrLevels.length-1) ? gotoGivenNode(objNode,arrLevels[i]-1,false,true) : gotoGivenNode(objNode,arrLevels[i]-1); 
        Delay(1000);
        while(objTree.VScroll.Pos <= intPos + 2){//whether under expanding or not
            Delay(1000);
            objTree.Refresh();
            if(i == arrLevels.length-1 || counter>20) break;//last node or timeout exit loop
            counter++;
        }
    }
    return objNode;
}

//-------------------------------------------------------------------------------------
//Function Name : gotoGivenNode
//Author        : Alan.Yang
//Create Date   : June 9, 2015
//Last Modify   : 
//Description   : goto and expand the specified node
//Parameter     : [IN]objNode -- the object Tree or Node
//Parameter     : [IN]level -- the level of node such as 6
//Parameter     : [IN]isFirstNode -- optional
//Parameter     : [IN]isLastNode -- optional
//Return        : Object, the specified node
//-------------------------------------------------------------------------------------
function gotoGivenNode(objNode, level, isFirstNode, isLastNode){
    isFirstNode = arguments[2] == undefined ? false : isFirstNode;
    isLastNode = arguments[3] == undefined ? false : isLastNode;
    try
    {
        if(isFirstNode) objNode = objNode.wItems.Item(0);
        if(objNode.Items != null && objNode.Items.Count>0){//when exists child node
            objNode = objNode.Items.Item(level);
        }
        else{
            Log.Error("Not exists any child nodes in Parent Node "+level +",stop finding.");
        }
        if(isLastNode){
            objNode.Click();//the last node needn't expanded.
            Log.Message("Select node:"+objNode.Text);
        }
        else{
            objNode.DblClick();
            Log.Message("Expanding node: "+objNode.Text);
        }
    }
    catch (e)
    {
        Log.Error(e.message);
    }
    return objNode;
}

//-------------------------------------------------------------------------------------
//Function Name : HashMap
//Author        : Alan.Yang
//Create Date   : September 21, 2015
//Last Modify   : 
//Description   : define a HashMap(keys,values) Class
//Parameter     : 
//Return        : object
//-------------------------------------------------------------------------------------
function HashMap(){
    //member's variables
    this.arrKey = new Array();//keys
    this.arrValue = new Array();//values
    
    //member's functions
    this.size = function(){
        return(this.arrKey.length);
    }
    
    this.intContainsKey = function(key){
        var index = -1;
        for(var i=0; i<this.arrKey.length; i++){
            if(this.arrKey[i] == key){
                index = i;
                break;
            }
        }
        return index;
    }
    
    this.put = function(key,value){
        var index = this.intContainsKey(key);
        if(index == -1){
            this.arrKey.push(key);
            this.arrValue.push(value);
        }else{
            this.arrValue[index] = value;
        }
    }
    
    this.get = function(key){
        var index = this.intContainsKey(key);
        if(index == -1){
            return null;
        }else{
            return this.arrValue[index];
        }
    }
    /*  you can expand more functions when you needed  */
}


