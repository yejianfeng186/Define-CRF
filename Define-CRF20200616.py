import pyocr;
# import importlib;
import sys;
import time;
# import xml.dom.minidom as minidom;
import os.path;
import pandas as pd;
import getpass;
import json;
import re;
# import numpy as np;
# import  multiprocessing;

from pdfminer.pdfparser import  PDFParser,PDFDocument;
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter;
from pdfminer.converter import PDFPageAggregator;
from pdfminer.layout import LTTextBoxHorizontal,LAParams, LTTextBox, LTTextLine;
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed;

# pd.set_option('display.max_columns', 10);

def parseCRF(strFilePath,strFileName):

    # print("Start parse CRF");
    # fp = open(strSysPath+"/CRF/"+strFileName,'rb');
    fp = open(strFilePath+"\\"+strFileName,'rb');
    
    parser = PDFParser(fp);

    doc = PDFDocument();

    parser.set_document(doc);
    doc.set_parser(parser);

    doc.initialize()
 

    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:

        rsrcmgr = PDFResourceManager();

        laparams = LAParams();
        device = PDFPageAggregator(rsrcmgr,laparams=laparams);

        interpreter = PDFPageInterpreter(rsrcmgr,device);

        count=0;
        listCRF=list();
        for page in doc.get_pages():
            interpreter.process_page(page);

            layout = device.get_result();
            # strForm="";
            
            objPageQuestions=createPageQuestions(layout,count+1);

            listCRF.extend(objPageQuestions);

            count=count+1;
            # if count >10:
            #     break;
                
        df=pd.DataFrame(listCRF);
        # print("End parse CRF");
        return df;
        
def createPageQuestions(objPage,intPageNum):
    listPage=list();
    intOrder=0;
    strForm="";
    for objQuestionGroup in objPage:
        if(isinstance(objQuestionGroup,LTTextBoxHorizontal)):
            
            for objQuestion in objQuestionGroup:
                dictQuestion=dict();
                strQuestion=objQuestion.get_text()[0:-1];
                # print(strQuestion.encode("GBK", 'ignore'));
                if "Blank CRF" in strQuestion:
                    pass;
                elif "only)" in strQuestion:
                    pass;
                elif "Project Name:" in strQuestion:
                    pass;
                elif "Form:" in strQuestion:
                    strForm=strQuestion.lstrip("Form:");
                    strForm=strForm.strip();
                elif "Generated On:" in strQuestion:
                    pass;
                elif objQuestion.y0<50:
                    pass;
                else:
                    intOrder=intOrder+1;
                    if objQuestion.x0 <=200:
                        dictQuestion['class']="question";
                    else:
                        continue;
                        dictQuestion['class']="codedata";

                    dictQuestion['pagenumber']=intPageNum;
                    dictQuestion['form']=strForm;
                    dictQuestion['x']=round(objQuestion.x0,3);
                    dictQuestion['y']=round(objQuestion.y0,3);
                    dictQuestion['width']=round(objQuestion.width,3);
                    dictQuestion['height']=round(objQuestion.height,3);
                    dictQuestion['pretext']=strQuestion.strip();
                    dictQuestion['order']=intOrder;

                    listPage.append(dictQuestion);
    
    return listPage;
    
def extractStudyFile(strFileName):
    
    dfForm=pd.read_excel(strFileName,sheet_name='Forms');
    dfFields=pd.read_excel(strFileName,sheet_name='Fields');
    dfDictionary=pd.read_excel(strFileName,sheet_name='DataDictionaryEntries');
    
    dfSubForm=dfForm.loc[:,["OID","DraftFormName"]];
    dfSubFields=dfFields.loc[:,["FormOID","DraftFieldName","PreText","Ordinal","DataDictionaryName"]];
    dfSubDictionary=dfDictionary.loc[:,["DataDictionaryName","UserDataString","Ordinal"]];
    
    dfSubFormNotNull=dfSubForm[dfSubForm["OID"].notnull()];
    dfSubFieldsNotNull=dfSubFields[dfSubFields["FormOID"].notnull()];
    dfSubDictionaryNotNull=dfSubDictionary[dfSubDictionary["DataDictionaryName"].notnull()];

    dfSubFieldsRename=dfSubFieldsNotNull.rename(columns={"DraftFieldName":"CDASH","DataDictionaryName":"codename"});
    dfSubDictionaryRename=dfSubDictionaryNotNull.rename(columns={"UserDataString":"PreText","DataDictionaryName":"codename"});
    
    dfSubFieldsRename["DraftDomain"]=dfSubFieldsRename["FormOID"].apply(lambda x: x.split("_")[0]);
    
    dfSub=pd.merge(dfSubFormNotNull,dfSubFieldsRename,left_on="OID",right_on="FormOID",how="right");
    dfSubFormRename=dfSub.rename(columns={"DraftFormName":"form"});
    dfDraftFormNameOnly=dfSubFormRename.loc[:,["codename","form","CDASH"]].copy();
    dfSubDictionaryForm=pd.merge(dfSubDictionaryRename,dfDraftFormNameOnly,left_on="codename",right_on="codename",how="left").copy();

    dfSubFormRename["PreText"]=dfSubFormRename["PreText"].replace(["<i>","</i>","<i/>","<b>","</b>","<br>","</br>","<br/>","\n"],"",regex=True);
    
    return dfSubFormRename,dfSubDictionaryForm;

def parseSDS(strFilePath,strFileName):

    # dfSheet=pd.read_excel(strFilePath+"\\"+strFileName,["Forms","Fields"]);
    # dfForm=dfSheet['Forms'];
    # dfFields=dfSheet['Fields'];

    dictMetaDataSds=pd.read_excel(strFilePath+"\\"+strFileName,None);
    dfMatrix=pd.DataFrame();
    for key in dictMetaDataSds:
        dfSheet=dictMetaDataSds[key].rename(columns=lambda x: x.strip())
        if key=="Forms":
            dfForm=dfSheet;
        elif key=="Fields":
            dfFields=dfSheet;
        elif key=="Folders":
            dfFolders=dfSheet;

        elif re.match("Matrix[0-9]+",key,re.I):
            strColunmName=dfSheet.columns[0]
            dfSheetRename=dfSheet.rename(columns={strColunmName:"CRFDS"})
            dfSheetNone=dfSheetRename[dfSheetRename.drop(columns=["CRFDS"]).any(axis='columns')];
            # print(dfSheetRename.notnull().any(axis='columns'))
            if dfMatrix.empty:
                dfMatrix=dfSheetNone;
            else:
                dfMatrix=pd.concat([dfMatrix,dfSheetNone]);
    dfMatrixSort=dfMatrix.sort_values(by=["CRFDS"],ascending=[True]);

    # dfFoldersRename=dfFolders.loc[:,["OID","FolderName"]].rename(columns={"OID":"CRFVISID","FolderName":"CRFVISIT"});
    dfFoldersRename=dfFolders.loc[:,["OID","FolderName","Ordinal"]].rename(columns={"OID":"CRFVISID","FolderName":"CRFVISIT","Ordinal":"CRFVISOD"});
   
    dfMatrixAgg=dfMatrixSort.groupby("CRFDS").apply(lambda x: x.any()).drop(columns=["CRFDS","Subject"]);
    # dfMatrixAgg.to_csv("aaa.csv");
    dfMatrixAgg["CRFVISID"]=dfMatrixAgg.apply(lambda x: x[x].index.str.cat(sep=','),axis=1);
    dfMatrixKeep=dfMatrixAgg["CRFVISID"].reset_index();
    
    dfMatrixUnstack=dfMatrixKeep.drop("CRFVISID", axis=1).join(dfMatrixKeep["CRFVISID"].str.split(",",expand=True).stack().reset_index(level=1, drop=True).rename("CRFVISID"));
    dfMatrixMerge=pd.merge(dfMatrixUnstack,dfFoldersRename,on="CRFVISID",how="inner")
    dfMatrixFinal=dfMatrixMerge.groupby(["CRFDS"]).agg(list)
    # dfMatrixAgg=dfMatrixSort.groupby("CRFDS").apply(lambda x: x.any()).drop(columns=["CRFDS"]);
    
    # dfMatrixAgg["CRFVISIT"]=dfMatrixAgg.apply(lambda x: x[x].index.tolist(),axis=1);
    # dfMatrixFinal=dfMatrixAgg["CRFVISIT"].reset_index();


    dfSubForm=dfForm.loc[:,["OID","DraftFormName"]];
    dfSubFields=dfFields.loc[:,["FormOID","DraftFieldName","PreText","Ordinal","DataDictionaryName","DraftFieldActive","DefaultValue","IsLog"]];
    
    dfSubFormNotNull=dfSubForm[dfSubForm["OID"].notnull()];
    dfSubFieldsNotNull=dfSubFields[dfSubFields["FormOID"].notnull()];
    # dfSubFieldsNotNull.to_csv("./aaa.csv");

    dfIsLogDef=dfSubFieldsNotNull[(dfSubFieldsNotNull["IsLog"]==True) & (dfSubFieldsNotNull["DefaultValue"].notnull())];
    # dfIsLogDef.to_csv("aaa.csv");
    dfIslogDefNodup=dfIsLogDef.loc[:,["FormOID","DefaultValue"]];
    # dfIslogDefNodup.to_csv("bb.csv");
    dfIslogDefNodup["COUNT"]=dfIslogDefNodup["DefaultValue"].apply(lambda x:x.count('|'));
    dfisFilterList=dfIslogDefNodup[dfIslogDefNodup["COUNT"]>1].drop(columns=["DefaultValue"]).drop_duplicates(["FormOID"],keep='last');
    # dfisFilterList.to_csv("aaa.csv");
    dfDeriveMerge=pd.merge(dfSubFieldsNotNull,dfisFilterList,on="FormOID",how="left");
    # dfDeriveMerge.to_csv("bbb.csv");
    dfDerive=dfDeriveMerge.apply(setRepeat, axis=1);

    # dfDerive.to_csv("ccc.csv");
    # print(dfSubFieldsNotNull["DraftFieldActive"]);
    # dfActive=dfSubFieldsNotNull[dfSubFieldsNotNull["DraftFieldActive"]];
    dfActive=dfDerive[dfDerive["DraftFieldActive"]];
    dfActive=dfActive.join(dfActive["DefaultValue"].str.split("|",expand=True).stack().reset_index(level=1, drop=True).rename("CRFVAL"));
    # dfActive.to_csv("./bbb.csv");

    dfSubFormRename=dfSubFormNotNull.rename(columns={"OID":"CRFDS","DraftFormName":"CRFDSLAB"});
    dfSubFieldsRename=dfActive.rename(columns={"DraftFieldName":"CRFVAR","DataDictionaryName":"CRFDIC","PreText":"CRFDES"});

    dfSubFieldsRename["CRFDOM"]=dfSubFieldsRename["FormOID"].apply(lambda x: x.split("_")[0]);
    dfSubFormVisit=pd.merge(dfSubFormRename,dfMatrixFinal,on="CRFDS",how="left");
    dfFormField=pd.merge(dfSubFormVisit,dfSubFieldsRename,left_on="CRFDS",right_on="FormOID",how="right");
    
    # dfSubFormField=dfFormField.loc[:,["CRFDIC","CRFDSLAB","CRFVAR"]].copy();

    dfFormField["CRFDES"]=dfFormField["CRFDES"].replace(["<i>","</i>","<i/>","<b>","</b>","<br>","</br>","<br/>","\n"],"",regex=True).str.strip();
    # dfFormAgg=dfFormField.fillna("").groupby(["CRFDS","CRFDSLAB","FormOID","CRFDES","CRFDIC","CRFDOM"],as_index=False).agg(";".join);
    # dfFormField.to_csv("aaa.csv");
    dfDropVar=dfFormField.drop(columns=["DraftFieldActive","DefaultValue"]);
    dfNotNull=dfDropVar[dfDropVar["CRFDES"].notnull()];
    # dfDropVar.to_csv("./ccc.csv");
    # print(dfFormField);



    return dfNotNull;

def setRepeat(val):

    if val["COUNT"]>1 and val["IsLog"]==True:
        if pd.isnull(val["DefaultValue"]):
            val["DefaultValue"]="|"*int(val["COUNT"]);
        elif val["DefaultValue"].count('|')==0:
            val["DefaultValue"]=(val["DefaultValue"]+"|")*int(val["COUNT"]);
    # if val["COUNT"]>1 and pd.isnull(val["DefaultValue"]) and val["IsLog"]==True:
    #     val["DefaultValue"]="|"*int(val["COUNT"]);
    # else if val["COUNT"]>1 and pd.isnull(val["DefaultValue"]) and val["IsLog"]==True:
    #     val["DefaultValue"]=(val["DefaultValue"]+"|")*int(val["COUNT"]);
    return val;

def parseSdsDic(strFilePath,strFileName):

    dfCrfDic=pd.read_excel(strFilePath+"\\"+strFileName,sheet_name='DataDictionaryEntries');
    dfSubCrfDic=dfCrfDic.loc[:,["DataDictionaryName","UserDataString","CodedData"]];

    dfSubCrfDicNotNull=dfSubCrfDic[dfSubCrfDic["DataDictionaryName"].notnull()];

    dfSubCrfDicRename=dfSubCrfDicNotNull.rename(columns={"UserDataString":"CRFCODELAB","DataDictionaryName":"CRFDIC","CodedData":"CRFCODE"});
    # dfSubCrfDicRename.to_csv("./a.csv");
    dfSubCrfDicFinal=dfSubCrfDicRename.where(dfSubCrfDicRename.notnull(), None);
    # dfSubCrfDicFinal.to_csv("./b.csv");
    return dfSubCrfDicFinal;
    # print("Dictionary finished Loading");

def parseIgCdash(strFilePath,strFileName):

    fileIgCdash = open(strFilePath+"\\std\\"+strFileName,'r',encoding='utf-8');
    jsonIgCdash = json.load(fileIgCdash);
    dfIgCdash=pd.DataFrame(jsonIgCdash);
    return dfIgCdash;

    
def parseSanofiCdash(strFilePath,strFileName):

    fileSanofiCdash = open(strFilePath+"\\std\\"+strFileName,'r',encoding='utf-8');
    jsonSanofiCdash = json.load(fileSanofiCdash);
    dfSanofiCdash=pd.DataFrame(jsonSanofiCdash);
    return dfSanofiCdash;

def parseIgSdtm(strFilePath,strFileName):

    fileIgSdtm = open(strFilePath+"\\std\\"+strFileName,'r',encoding='utf-8');
    jsonIgSdtm = json.load(fileIgSdtm);
    dfIgSdtm=pd.DataFrame(jsonIgSdtm);
    return dfIgSdtm;

def parseSanofiSdtm(strFilePath,strFileName):
    fileSanofiSdtm = open(strFilePath+"\\std\\"+strFileName,'r',encoding='utf-8');
    jsonSanofiSdtm = json.load(fileSanofiSdtm);
    dfSanofiSdtm=pd.DataFrame(jsonSanofiSdtm);
    return dfSanofiSdtm;

def parseSanofiSdtmDom(strFilePath,strFileName):
    fileSanofiSdtmDom = open(strFilePath+"\\std\\"+strFileName,'r',encoding='utf-8');
    jsonSanofiSdtmDom = json.load(fileSanofiSdtmDom);
    dfSanofiSdtmDom=pd.DataFrame(jsonSanofiSdtmDom);
    return dfSanofiSdtmDom;
    
# def mapSDS(dfCRF,dfSDS,dfDictionary,dfCDASH,dfSanofiSDTM,dfSanofiMapping):
# def mapSDS(dfCRF,dfSDS,dfSdsDic,dfIgCdash,dfSanofiCdash,dfIgSdtm,dfSanofiSdtm,dfSanofiSdtmDom):
def mapSDS(dfCRF,dfSDS,dfSdsDic,dfIgCdash,dfSanofiCdash,dfIgSdtm,dfSanofiSdtm):
    dfCrfQue=dfCRF[dfCRF["class"]=="question"].copy();
    dfCrfQue["CDASHTYPE"]="SANOFI";
    dfNodupVeriable=mergeQuestion(dfCrfQue,dfSDS).copy();
    # dfNodupVeriable.to_csv("dfNodupVeriable.csv");
    # dfAllSet=pd.concat([dfCrfSdsMap,dfNodupVeriable],ignore_index=True,sort=False);
    dfStdAllSet=dfNodupVeriable.loc[:,["PAGENUMBER","ORDER","CLASS","X","Y","HEIGHT","WIDTH","CRFDS","CRFVISIT","CRFVISID","CRFVISOD","CRFDSLAB","CRFVAR","CRFDES","CRFDIC","CRFDOM","CDASHTYPE","CRFVAL"
]];
    # dfStdAllSet["MCDASH"]= dfStdAllSet["CRFVAR"].str.replace(r'[\(\)\d]+', '');
    # dfSdsDicAgg=dfSdsDic[["CRFDIC","CRFCODE"]].groupby("CRFDIC",as_index=False).agg(";".join);
    dfSdsDicAgg=dfSdsDic[["CRFDIC","CRFCODE"]].groupby("CRFDIC",as_index=False).agg(list);
    dfMapSdsDic=pd.merge(dfStdAllSet,dfSdsDicAgg,on="CRFDIC",how="left");

    dfMapIgCdash=pd.merge(dfMapSdsDic,dfIgCdash,left_on="CRFVAR",right_on="CDASHVAR",how="left");
    dfSanofiCdashRename=dfSanofiCdash.rename(columns={"CDASHVAR":"SANOFICDASHVAR","SDTMDOM":"SANOFISDTMDOM","SDTMVAR":"SANOFISDTMVAR"});
    dfMapSanofiCdash=pd.merge(dfMapIgCdash,dfSanofiCdashRename,left_on="CRFVAR",right_on="SANOFICDASHVAR",how="left");
    dfMapCdashType=dfMapSanofiCdash.apply(setCdashType, axis=1).drop(columns=["SANOFICDASHVAR","SANOFISDTMDOM","SANOFISDTMVAR"]).drop(columns=["CDASHQUE"]);
    # dfMapCdashType.to_csv("aaa.csv");
    dfIgSdtmRename=dfIgSdtm.rename(columns={"SDTMDOM":"IGSDTMDOM","SDTMVAR":"IGSDTMVAR"});
    dfMapIgSdtm=pd.merge(dfMapCdashType,dfIgSdtmRename,left_on="CRFVAR",right_on="IGSDTMVAR",how="left");
    
    dfSanofiSdtmRename=dfSanofiSdtm.rename(columns={"SDTMDOM":"SANOFISDTMDOM","SDTMVAR":"SANOFISDTMVAR"});
    dfMapSanofiSdtm=pd.merge(dfMapIgSdtm,dfSanofiSdtmRename,left_on="CRFVAR",right_on="SANOFISDTMVAR",how="left");

    dfMapSdtmType=dfMapSanofiSdtm.apply(setSdtmType, axis=1).drop(columns=["IGSDTMDOM","IGSDTMVAR","SANOFISDTMDOM","SANOFISDTMVAR"]);
    dfAllNotNull=dfMapSdtmType[dfMapSdtmType["PAGENUMBER"].notnull()];
    dfAllSort=dfAllNotNull.sort_values(by=["PAGENUMBER","Y","X"],ascending=[True,False,True]);
    # dfAllSort.to_csv("ddd.csv");
    # dfAllSort.fillna("");
    dfAllSortRep=dfAllSort.where(dfAllSort.notnull(), None);
    return dfAllSortRep;
    
def setCdashType(var):

    if not isinstance(var["SANOFISDTMVAR"],list):
        var["SANOFISDTMVAR"]=[];

    if pd.notnull(var.CDASHVAR):
        var.CDASHTYPE="IG";
        # if pd.notnull(var.SDTMVAR):
        if len(var["SDTMVAR"])>0:
            var["SDTMTYPE"]="IG";
            var["SDTMDOM"]=var.CDASHDOM;
            
            # if pd.notnull(var.SANOFISDTMVAR):
            if len(var["SANOFISDTMVAR"])>0:
                var.SDTMVAR=var.SANOFISDTMVAR;
                var["SDTMDOM"]=var.SANOFISDTMDOM;
                
        elif len(var["SANOFISDTMVAR"])>0:
            var["SDTMTYPE"]="SANOFI";
            var.SDTMVAR=var.SANOFISDTMVAR;
            var["SDTMDOM"]=var.SANOFISDTMDOM;

    elif pd.notnull(var.SANOFICDASHVAR):
    # elif var["SANOFICDASHVAR"].length>0:
        var.CDASHDOM=[var.CRFDOM];
        var.CDASHVAR=var.CRFVAR;
        # var.CDASHQUE=var.CRFDES;
        
        # if pd.notnull(var.SANOFISDTMVAR):
        if len(var["SANOFISDTMVAR"])>0:
            var["SDTMTYPE"]="SANOFI";
            var.SDTMVAR=var.SANOFISDTMVAR;
            var["SDTMDOM"]=var.SANOFISDTMDOM;
            
    elif pd.notnull(var.CRFVAR):
        var.CDASHDOM=[var.CRFDOM];
        var.CDASHVAR=var.CRFVAR;
            # var.CDASHQUE=var.CRFDES;
    else:
        var.CDASHDOM=[var.CRFDOM];
        var.CDASHVAR=var.CRFVAR;
        
    return var;
    
def setSdtmType(var):
    if not isinstance(var["SDTMVAR"],list):
        if pd.notnull(var.IGSDTMVAR):
            var.SDTMTYPE="IG";
            # if pd.isnull(var.SDTMVAR):
            var.SDTMDOM=var.IGSDTMDOM;
            var.SDTMVAR=[var.IGSDTMVAR];
                
        elif pd.notnull(var.SANOFISDTMVAR):
            if pd.isnull(var.SDTMTYPE):
                var.SDTMTYPE="SANOFI";
            var.SDTMDOM=var.SANOFISDTMDOM;
            var.SDTMVAR=[var.SANOFISDTMVAR];
    return var;

def offsetCoordinate(var):
    if var.OBJCOUNT >1 :
        var.X=var.X+var.OBJRANK;
        var.Y=var.Y+var.OBJRANK;
    return var;
    
def mergeQuestion(dfCrf,dfSds):
    
    dfTargetSDS=dfSds[dfSds.CRFDSLAB.isin(dfCrf.form)].copy();
    dfTargetCrf=dfCrf.reset_index(level=0).copy();
    dfSDS=dfTargetSDS.copy();
    # dfCrf.to_csv("a.csv");
    # dfTargetCrf.to_csv("b.csv");

    for i,i_row in dfTargetSDS.iterrows():
        # seriousTargetVeriable=dfCrf[i_row.CRFDSLAB==dfCrf.form];
        seriousTargetVeriable=dfTargetCrf[i_row.CRFDSLAB==dfTargetCrf.form];
        dfSDS.loc[i,"ORDER"]=0;
        dfSDS.loc[i,"X"]=-100;
        dfSDS.loc[i,"Y"]=760;
        dfSDS.loc[i,"CLASS"]="unassigned";
        dfSDS.loc[i,"HEIGHT"]=15;
        dfSDS.loc[i,"WIDTH"]=100;
   
        for j,j_row in seriousTargetVeriable.iterrows():
            dfSDS.loc[i,"PAGENUMBER"]=j_row.pagenumber;
            dfSDS.loc[i,"CDASHTYPE"]=j_row.CDASHTYPE;
            if i_row.CRFDES == j_row.pretext:
                dfSDS.loc[i,"ORDER"]=j_row.order;
                dfSDS.loc[i,"CLASS"]=j_row["class"];
                dfSDS.loc[i,"X"]=j_row.x;
                dfSDS.loc[i,"Y"]=j_row.y;
                dfSDS.loc[i,"HEIGHT"]=j_row.height;
                dfSDS.loc[i,"WIDTH"]=j_row.width;
                
                dfTargetCrf.drop(j,axis=0,inplace=True);
                break;

    dfSdsMap=dfSDS[dfSDS["CLASS"]=="question"];
    dfSdsNotMap=dfSDS[(dfSDS["CLASS"]!="question")];
    # print("START");
    for m,m_row in dfSdsNotMap.iterrows():
        # print(m_row);
        seriousTargetVeriable=dfTargetCrf.copy();
        for n,n_row in seriousTargetVeriable.iterrows():
            # print();
            if (m_row["CRFDES"] in n_row["pretext"]) or (n_row["pretext"] in m_row["CRFDES"]) :
                m_row["PAGENUMBER"]=n_row["pagenumber"];
                m_row["CDASHTYPE"]=n_row["CDASHTYPE"];
                m_row["ORDER"]=n_row["order"];
                m_row["CLASS"]=n_row["class"];
                m_row["X"]=n_row["x"];
                m_row["Y"]=n_row["y"];
                m_row["HEIGHT"]=n_row["height"];
                m_row["WIDTH"]=n_row["width"];
                dfTargetCrf.drop(n,axis=0,inplace=True);
                break;
        # print(m);
    # dfSdsNotMap.to_csv("c.csv");
    # print(3);
    dfSdsConcat=pd.concat([dfSdsMap,dfSdsNotMap],ignore_index=True,sort=False);
    
    dfSdsSort=dfSdsConcat.sort_values(by=["PAGENUMBER","Y","X"],ascending=[True,False,True]);
    
    dfSdsFinal=dfSdsSort.reset_index(level=0,drop=True);
    # dfSdsFinal.to_csv("d.csv");
    # dfTargetCrf.to_csv("e.csv");
    # dfNodupSDS=dfSDS.drop_duplicates(["CRFDSLAB","CRFDES"],keep='last');
    # dfNodupSDS.to_csv("b.csv");
    return dfSdsFinal;

def mergeCodedata(dfValueNonMap,dfDictionary):
    dfNewValueNonMap=dfValueNonMap.copy();

    for i,row in dfValueNonMap.iterrows():
        seriousTargetForm=dfDictionary[dfDictionary.form==row.form];
        
        for j,rowDictionary in seriousTargetForm.iterrows():
            if  row.pretext in rowDictionary.PreText:
                
                dfNewValueNonMap.loc[i,"CDASH"]=seriousTargetForm.loc[j,"CDASH"];
                dfNewValueNonMap.loc[i,"codename"]=seriousTargetForm.loc[j,"codename"];
                dfNewValueNonMap.loc[i,"PreText"]=seriousTargetForm.loc[j,"PreText"];
                break;

    dfNodupValue=dfNewValueNonMap.drop_duplicates(["class","form","PreText"],keep='last');

    return dfNodupValue;

if __name__ == '__main__':

    strSysPath=os.path.dirname(os.path.abspath(sys.argv[0]));
    # multiprocessing.freeze_support();
    # print(strSysPath);
    # input("Enter:");
    fileProfile = open(strSysPath+"\\init.json",'r',encoding='utf-8');
    # strSysPath=os.getcwd();
    strCurrentDatetime=time.strftime("%Y%m%dT%H%M%S", time.localtime());

    jsonInfo = json.load(fileProfile);
    strCRFFullName=jsonInfo["CRF"];
    strSDSFullName=jsonInfo["SDS"];
    strShowVarType   =jsonInfo["ShowVarType"];
    strShowDSType   =jsonInfo["ShowDSType"];
    strAdmin=getpass.getuser();

    if len(strCRFFullName) >0 :
        (strCRFPath, strCRFName) = os.path.split(strCRFFullName);
        
    if len(strCRFFullName) >0 :
        (strSDSPath, strSDSName) = os.path.split(strSDSFullName);
    
    # print(strCRFPath);
    # print(strSDSPath);
    # print(strCRFName);
    # print(strSDSName);
    
    if len(strCRFPath) ==0 :
        strCRFPath=".";
    
    if len(strSDSPath) ==0 :
        strSDSPath=".";
    
    if fileProfile :
        time1 = time.time();
        # pool = multiprocessing.Pool();
        fileMapping =open(strSysPath+"\\std\\METADATAMAPPING.json",'r',encoding='utf-8');
        dfMapping = json.load(fileMapping);
        # dfMapping=pd.DataFrame(jsonMapping);

        # dfSDS=pool.apply_async(func=parseSDS, args=(strSDSPath,strSDSName,)).get();
        dfSDS=parseSDS(strSDSPath,strSDSName);
        # print(time.time());
        dfSdsDic=parseSdsDic(strSDSPath,strSDSName);
        
        # dfSdsDic=pool.apply_async(func=parseSdsDic, args= (strSDSPath,strSDSName ,)).get();
        # print(time.time());
        # dfSanofiSdtmDom=pool.apply_async(func=parseSanofiSdtmDom, args=(strSysPath,"SANOFISDTMDOM.json",)).get();
        # print(time.time());
        dfCRF=parseCRF(strCRFPath,strCRFName);
        # dfCRF.to_csv("aaa.csv");
        
        # dfCRF=pool.apply_async(func=parseCRF, args=(strCRFPath,strCRFName ,)).get();
        # print(time.time());
        # pool.close();
        # pool.join();
        # print(time.time());
        dfIgCdash=pd.DataFrame(dfMapping["IGCDASH"]);
        dfSanofiCdash=pd.DataFrame(dfMapping["SANOFICDASH"]);
        dfIgSdtm=pd.DataFrame(dfMapping["IGSDTM"]);
        dfSanofiSdtm=pd.DataFrame(dfMapping["SANOFISDTM"]);
  
        # dfIgCdash=pool.apply_async(func=parseIgCdash, args=(strSysPath,"IGCDASH.json",)).get();
        # dfSanofiCdash=pool.apply_async(func=parseSanofiCdash, args=(strSysPath,"SANOFICDASH.json",)).get();
        # dfIgSdtm=pool.apply_async(func=parseIgSdtm, args=(strSysPath,"IGSDTM.json",)).get();
        # dfSanofiSdtm=pool.apply_async(func=parseSanofiSdtm, args=(strSysPath,"SANOFISDTM.json",)).get();

        # dfACRF=mapSDS(dfCRF,dfSDS,dfSdsDic,dfIgCdash,dfSanofiCdash,dfIgSdtm,dfSanofiSdtm,dfSanofiSdtmDom);
        dfACRF=mapSDS(dfCRF,dfSDS,dfSdsDic,dfIgCdash,dfSanofiCdash,dfIgSdtm,dfSanofiSdtm);
        if len(strShowDSType) ==0 :
            strShowDSType="CDASH";
            
        if len(strShowVarType) ==0 :
            strShowVarType="BOTH";
        # createAnnotation(dfACRF,strCRFName,strShowVarType);
        # createAnnotation(dfACRF,strCRFPath,strCRFName,strCurrentDatetime,strShowDSType,strShowVarType);
        
        strFilename=strCRFName.split(".")[0];
        dictMeta={"NAME":strAdmin,"VERSION":"1.0.0","CREATED":strCurrentDatetime};
        dictMeta["MCRF"]=dfACRF.to_dict(orient="records");

        with open(strCRFPath+"\\"+strFilename+strCurrentDatetime+".mcrf",'w') as f:
            # print(dfACRF.to_json(orient='records'));
            f.write(json.dumps(dictMeta));
        dfACRF.to_csv(strCRFPath+"\\"+strFilename+strCurrentDatetime+".csv");
        # dictMeta.to_json(strCRFPath+"/"+strFilename+strCurrentDatetime+".json",orient='records');
        # dfACRF.to_json(strSysPath+"/data/"+strFilename+strCurrentDatetime+".json",orient='records');
        time2 = time.time();
        print(strCRFPath+"\\"+strFilename+strCurrentDatetime+".mcrf");

        # print("Expended:",time2-time1);