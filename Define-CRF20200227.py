import pyocr;
# import importlib;
import sys;
import time;
import xml.dom.minidom as minidom;
import os.path;
import pandas as pd;
import getpass;
import json;
import re;
# import numpy as np;
import  multiprocessing;

from pdfminer.pdfparser import  PDFParser,PDFDocument;
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter;
from pdfminer.converter import PDFPageAggregator;
from pdfminer.layout import LTTextBoxHorizontal,LAParams, LTTextBox, LTTextLine;
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed;

pd.set_option('display.max_columns', 10);

def parseCRF(strFilePath,strFileName):

    # print("Start parse CRF");
    # fp = open(strSysPath+"/CRF/"+strFileName,'rb');
    fp = open(strFilePath+"/"+strFileName,'rb');
    
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
            # if count >20:
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
                    dictQuestion['pagenumber']=intPageNum;
                    dictQuestion['form']=strForm;
                    dictQuestion['x']=round(objQuestion.x0,3);
                    dictQuestion['y']=round(objQuestion.y0,3);
                    dictQuestion['width']=round(objQuestion.width,3);
                    dictQuestion['height']=round(objQuestion.height,3);
                    dictQuestion['pretext']=strQuestion;
                    dictQuestion['order']=intOrder;
                    # dictQuestion['SDTM']="SDTMTEST";
                    
                    if objQuestion.x0 <=200:
                        dictQuestion['class']="question";
                    else:
                        dictQuestion['class']="codedata";
                        
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

    dfSheet=pd.read_excel(strFilePath+"\\"+strFileName,["Forms","Fields"]);
    dfForm=dfSheet['Forms'];
    dfFields=dfSheet['Fields'];
    
    dfSubForm=dfForm.loc[:,["OID","DraftFormName"]];
    dfSubFields=dfFields.loc[:,["FormOID","DraftFieldName","PreText","Ordinal","DataDictionaryName"]];
    
    dfSubFormNotNull=dfSubForm[dfSubForm["OID"].notnull()];
    dfSubFieldsNotNull=dfSubFields[dfSubFields["FormOID"].notnull()];
    
    dfSubFormRename=dfSubFormNotNull.rename(columns={"OID":"CRFDS","DraftFormName":"CRFDSLAB"});
    dfSubFieldsRename=dfSubFieldsNotNull.rename(columns={"DraftFieldName":"CRFVAR","DataDictionaryName":"CRFDIC","PreText":"CRFDES"});

    dfSubFieldsRename["CRFDOM"]=dfSubFieldsRename["FormOID"].apply(lambda x: x.split("_")[0]);
    
    dfFormField=pd.merge(dfSubFormRename,dfSubFieldsRename,left_on="CRFDS",right_on="FormOID",how="right");
    
    # dfSubFormField=dfFormField.loc[:,["CRFDIC","CRFDSLAB","CRFVAR"]].copy();

    dfFormField["CRFDES"]=dfFormField["CRFDES"].replace(["<i>","</i>","<i/>","<b>","</b>","<br>","</br>","<br/>","\n"],"",regex=True);
    # dfFormAgg=dfFormField.fillna("").groupby(["CRFDS","CRFDSLAB","FormOID","CRFDES","CRFDIC","CRFDOM"],as_index=False).agg(";".join);
    # dfFormField.to_csv("aaa.csv");
    # dfFormAgg.to_csv("aaa.csv");
    return dfFormField;

def parseSdsDic(strFilePath,strFileName):

    dfCrfDic=pd.read_excel(strFilePath+"/"+strFileName,sheet_name='DataDictionaryEntries');
    dfSubCrfDic=dfCrfDic.loc[:,["DataDictionaryName","UserDataString","CodedData"]];

    dfSubCrfDicNotNull=dfSubCrfDic[dfSubCrfDic["DataDictionaryName"].notnull()];

    dfSubCrfDicRename=dfSubCrfDicNotNull.rename(columns={"UserDataString":"CRFCODELAB","DataDictionaryName":"CRFDIC","CodedData":"CRFCODE"});

    return dfSubCrfDicRename;
    # print("Dictionary finished Loading");

def parseIgCdash(strFilePath,strFileName):

    fileIgCdash = open(strFilePath+"/std/"+strFileName,'r',encoding='utf-8');
    jsonIgCdash = json.load(fileIgCdash);
    dfIgCdash=pd.DataFrame(jsonIgCdash);
    return dfIgCdash;

    
def parseSanofiCdash(strFilePath,strFileName):

    fileSanofiCdash = open(strFilePath+"/std/"+strFileName,'r',encoding='utf-8');
    jsonSanofiCdash = json.load(fileSanofiCdash);
    dfSanofiCdash=pd.DataFrame(jsonSanofiCdash);
    return dfSanofiCdash;

def parseIgSdtm(strFilePath,strFileName):

    fileIgSdtm = open(strFilePath+"/std/"+strFileName,'r',encoding='utf-8');
    jsonIgSdtm = json.load(fileIgSdtm);
    dfIgSdtm=pd.DataFrame(jsonIgSdtm);
    return dfIgSdtm;

def parseSanofiSdtm(strFilePath,strFileName):
    fileSanofiSdtm = open(strFilePath+"/std/"+strFileName,'r',encoding='utf-8');
    jsonSanofiSdtm = json.load(fileSanofiSdtm);
    dfSanofiSdtm=pd.DataFrame(jsonSanofiSdtm);
    return dfSanofiSdtm;

def parseSanofiSdtmDom(strFilePath,strFileName):
    fileSanofiSdtmDom = open(strFilePath+"/std/"+strFileName,'r',encoding='utf-8');
    jsonSanofiSdtmDom = json.load(fileSanofiSdtmDom);
    dfSanofiSdtmDom=pd.DataFrame(jsonSanofiSdtmDom);
    return dfSanofiSdtmDom;
    
# def mapSDS(dfCRF,dfSDS,dfDictionary,dfCDASH,dfSanofiSDTM,dfSanofiMapping):
def mapSDS(dfCRF,dfSDS,dfSdsDic,dfIgCdash,dfSanofiCdash,dfIgSdtm,dfSanofiSdtm,dfSanofiSdtmDom):
    dfCrfQue=dfCRF[dfCRF["class"]=="question"].copy();
    dfCrfQue["CDASHTYPE"]="SANOFI";
    dfCrfSds=pd.merge(dfCrfQue,dfSDS,left_on=["form","pretext"],right_on=["CRFDSLAB","CRFDES"],how="left").copy();
    dfSdsCrf=pd.merge(dfCrfQue,dfSDS,left_on=["form","pretext"],right_on=["CRFDSLAB","CRFDES"],how="right").copy();
    
    dfCrfSdsMap=dfCrfSds[dfCrfSds["CRFVAR"].notnull()==True].copy();
    dfCrfSdsNoMap=dfCrfSds[dfCrfSds["CRFVAR"].isnull()==True].copy();
    dfSdsCrfNoMap=dfSdsCrf[dfSdsCrf["pretext"].isnull()==True].copy();
    
    # dfCrfSdsMap.to_csv("aaa.csv");
    # dfCrfSdsNoMap.to_csv("bbb.csv");
    # dfSdsCrfNoMap.to_csv("ccc.csv");
    # dfSdsNotInCrf=dfSDS[~(dfSDS.CRFDES.isin(dfSdsCrfMap.CRFDES) & dfSDS.CRFDSLAB.isin(dfSdsCrfMap.CRFDSLAB))];
    dfCrfSdsMap["PAGENUMBER"]=dfCrfSdsMap.pagenumber;
    dfCrfSdsMap["ORDER"]=dfCrfSdsMap.order;
    dfCrfSdsMap["CLASS"]=dfCrfSdsMap["class"];
    dfCrfSdsMap["X"]=dfCrfSdsMap.x;
    dfCrfSdsMap["Y"]=dfCrfSdsMap.y;
    dfCrfSdsMap["HEIGHT"]=dfCrfSdsMap.height;
    dfCrfSdsMap["WIDTH"]=dfCrfSdsMap.width;
    
    dfNodupVeriable=mergeQuestion(dfCrfSdsNoMap,dfSdsCrfNoMap).copy();

    # dfNodupVeriable.to_csv("dfNodupVeriable.csv");
    dfAllSet=pd.concat([dfCrfSdsMap,dfNodupVeriable],ignore_index=True,sort=False);
    dfStdAllSet=dfAllSet.loc[:,["PAGENUMBER","ORDER","CLASS","X","Y","HEIGHT","WIDTH","CRFDS","CRFDSLAB","CRFVAR","CRFDES","CRFDIC","CRFDOM","CDASHTYPE"
]];     
    # dfStdAllSet["MCDASH"]= dfStdAllSet["CRFVAR"].str.replace(r'[\(\)\d]+', '');
    dfSdsDicAgg=dfSdsDic[["CRFDIC","CRFCODE"]].groupby("CRFDIC",as_index=False).agg(";".join);
    dfMapSdsDic=pd.merge(dfStdAllSet,dfSdsDicAgg,on="CRFDIC",how="left");

    dfMapIgCdash=pd.merge(dfMapSdsDic,dfIgCdash,left_on="CRFVAR",right_on="CDASHVAR",how="left");

    dfSanofiCdashRename=dfSanofiCdash.rename(columns={"CDASHVAR":"SANOFICDASHVAR","SDTMDOM":"SANOFISDTMDOM","SDTMVAR":"SANOFISDTMVAR"});
    dfMapSanofiCdash=pd.merge(dfMapIgCdash,dfSanofiCdashRename,left_on="CRFVAR",right_on="SANOFICDASHVAR",how="left");
    
    dfMapCdashType=dfMapSanofiCdash.apply(setCdashType, axis=1).drop(columns=["SANOFICDASHVAR","SANOFISDTMDOM","SANOFISDTMVAR"]).drop(columns=["CDASHQUE"]);
    
    dfIgSdtmRename=dfIgSdtm.rename(columns={"SDTMDOM":"IGSDTMDOM","SDTMVAR":"IGSDTMVAR"});
    dfMapIgSdtm=pd.merge(dfMapCdashType,dfIgSdtmRename,left_on="CRFVAR",right_on="IGSDTMVAR",how="left");
    
    dfSanofiSdtmRename=dfSanofiSdtm.rename(columns={"SDTMDOM":"SANOFISDTMDOM","SDTMVAR":"SANOFISDTMVAR"});
    dfMapSanofiSdtm=pd.merge(dfMapIgSdtm,dfSanofiSdtmRename,left_on="CRFVAR",right_on="SANOFISDTMVAR",how="left");
    
    dfMapSdtmType=dfMapSanofiSdtm.apply(setSdtmType, axis=1).drop(columns=["IGSDTMDOM","IGSDTMVAR","SANOFISDTMDOM","SANOFISDTMVAR"]);
    dfSdtmDomSplit=dfMapSdtmType["SDTMDOM"].str.split(';', expand=True).stack().reset_index(drop=False).drop(columns=["level_1"]).rename(columns={0:"SDTMDOM"});
    
    # dfSdtmDomSplit=dfMapSdtmType["SDTMDOM"].str.split(';', expand=True).stack().reset_index(level=1,drop=True).rename("SDTMDOM");
    dfSdtmDomMergeDes=pd.merge(dfSdtmDomSplit,dfSanofiSdtmDom,on="SDTMDOM",how="left");
    
    dfSdtmDomMergeDes["DOMDES"]=dfSdtmDomMergeDes.apply(lambda x: x.SDTMDOM+"="+x.SDTMDOMLAB.upper() if pd.notnull(x.SDTMDOMLAB) else "" , axis=1);
    
    dfSdtmDomNoDup=dfSdtmDomMergeDes.drop_duplicates(keep="first");
    dfSdtmDomAgg=dfSdtmDomNoDup.fillna('').groupby("level_0",as_index=True).agg(";".join);
    dfSdtmDomJoin=dfMapSdtmType.drop(columns=["SDTMDOM"]).join(dfSdtmDomAgg);
    dfSdtmDomJoin["OBJCOUNT"]=dfSdtmDomJoin.groupby(["PAGENUMBER","X","Y"])["SDTMDOM"].transform("size");
    dfSdtmDomJoin["OBJRANK"]=dfSdtmDomJoin.groupby(["PAGENUMBER","X","Y"])["X"].rank(ascending=0,method='first');
    dfSdtmOffset=dfSdtmDomJoin.apply(offsetCoordinate, axis=1);
    # dfSdtmOffset.to_csv("ccc.csv");
    dfAllSort=dfSdtmOffset.sort_values(by=["PAGENUMBER","Y","ORDER","OBJRANK"],ascending=[True,False,True,True]);
    
    # dfAllSort.to_csv("ddd.csv");
    # dfAllSort.fillna("");
    dfAllSortRep=dfAllSort.where(dfAllSort.notnull(), None);
    return dfAllSortRep;
    
def setCdashType(var):
    if pd.notnull(var.CDASHVAR):
        var.CDASHTYPE="IG";
        if pd.notnull(var.SDTMVAR):
            var["SDTMTYPE"]="IG";
            var["SDTMDOM"]=var.CDASHDOM;
            
            if pd.notnull(var.SANOFISDTMVAR):
                var.SDTMVAR=var.SANOFISDTMVAR;
                var["SDTMDOM"]=var.SANOFISDTMDOM;
                
        elif pd.notnull(var.SANOFISDTMVAR):
            var["SDTMTYPE"]="SANOFI";
            var.SDTMVAR=var.SANOFISDTMVAR;
            var["SDTMDOM"]=var.SANOFISDTMDOM;
            
    elif pd.notnull(var.SANOFICDASHVAR):
        var.CDASHDOM=var.CRFDOM;
        var.CDASHVAR=var.CRFVAR;
        # var.CDASHQUE=var.CRFDES;
        
        if pd.notnull(var.SANOFISDTMVAR):
            var["SDTMTYPE"]="SANOFI";
            var.SDTMVAR=var.SANOFISDTMVAR;
            var["SDTMDOM"]=var.SANOFISDTMDOM;
            
    elif pd.notnull(var.CRFVAR):
        var.CDASHDOM=var.CRFDOM;
        var.CDASHVAR=var.CRFVAR;
            # var.CDASHQUE=var.CRFDES;
    else:
        var.CDASHDOM=var.CRFDOM;
        var.CDASHVAR=var.CRFVAR;
        
    return var;
    
def setSdtmType(var):
    if pd.notnull(var.IGSDTMVAR):
        var.SDTMTYPE="IG";
        # if pd.isnull(var.SDTMVAR):
        var.SDTMDOM=var.IGSDTMDOM;
        var.SDTMVAR=var.IGSDTMVAR;
            
    elif pd.notnull(var.SANOFISDTMVAR):
        if pd.isnull(var.SDTMTYPE):
            var.SDTMTYPE="SANOFI";
        var.SDTMDOM=var.SANOFISDTMDOM;
        var.SDTMVAR=var.SANOFISDTMVAR;
    return var;

def offsetCoordinate(var):
    if var.OBJCOUNT >1 :
        var.X=var.X+var.OBJRANK;
        var.Y=var.Y+var.OBJRANK;
    return var;
    
def mergeQuestion(dfCrfSdsNoMap,dfSdsCrfNoMap):
    
    dfTargetSDS=dfSdsCrfNoMap[dfSdsCrfNoMap.CRFDSLAB.isin(dfCrfSdsNoMap.form)];
    dfSDS=dfTargetSDS.copy();
    
    # dfCrfSdsNoMap.to_csv("a.csv");
    for i ,i_row in dfTargetSDS.iterrows():
        seriousTargetVeriable=dfCrfSdsNoMap[i_row.CRFDSLAB==dfCrfSdsNoMap.form];
        
        dfSDS.loc[i,"ORDER"]=0;
        dfSDS.loc[i,"X"]=-100;
        dfSDS.loc[i,"Y"]=760;
        dfSDS.loc[i,"CLASS"]="unassigned";
        dfSDS.loc[i,"HEIGHT"]=15;
        dfSDS.loc[i,"WIDTH"]=100;
        for j,j_row in seriousTargetVeriable.iterrows():

            dfSDS.loc[i,"PAGENUMBER"]=j_row.pagenumber;
            dfSDS.loc[i,"CDASHTYPE"]=j_row.CDASHTYPE;

            if (i_row.CRFDES in j_row.pretext) or (j_row.pretext in i_row.CRFDES) :
                dfSDS.loc[i,"ORDER"]=j_row.order;
                dfSDS.loc[i,"CLASS"]=j_row["class"];
                dfSDS.loc[i,"X"]=j_row.x;
                dfSDS.loc[i,"Y"]=j_row.y;
                dfSDS.loc[i,"HEIGHT"]=j_row.height;
                dfSDS.loc[i,"WIDTH"]=j_row.width;
                # dfSDS.loc[i,"CRFDIC"]=j_row.CRFDIC;
                
                dfCrfSdsNoMap=dfCrfSdsNoMap.drop(j);
                break;
            
    dfNodupSDS=dfSDS.drop_duplicates(["CRFDSLAB","CRFDES"],keep='last');
    # dfNodupSDS.to_csv("b.csv");
    return dfNodupSDS;
# def mergeQuestion2(dfVeriableNonMap,dfSubSDS):

    # dfNewSubSDS=dfSubSDS.copy();
    # dfTargetSDS=dfSubSDS[dfSubSDS.form.isin(dfVeriableNonMap.form)];

    # for i ,i_row in dfTargetSDS.iterrows():
        # seriousTargetVeriable=dfVeriableNonMap[i_row.form==dfVeriableNonMap.form];
        
        # dfNewSubSDS.loc[i,"order"]=0;
        # dfNewSubSDS.loc[i,"x"]=70;
        # dfNewSubSDS.loc[i,"y"]=700;
        # dfNewSubSDS.loc[i,"class"]="unassigned";
        # for j,j_row in seriousTargetVeriable.iterrows():

            # dfNewSubSDS.loc[i,"pagenumber"]=j_row.pagenumber;
            # dfNewSubSDS.loc[i,"CDASHTYPE"]=j_row.CDASHTYPE;
            
            # if (i_row.PreText in j_row.pretext) or (j_row.pretext in i_row.PreText) :
                # dfNewSubSDS.loc[i,"order"]=j_row.order;
                # dfNewSubSDS.loc[i,"class"]=j_row["class"];
                # dfNewSubSDS.loc[i,"x"]=j_row.x;
                # dfNewSubSDS.loc[i,"y"]=j_row.y;
                # dfNewSubSDS.loc[i,"height"]=j_row.height;
                # dfNewSubSDS.loc[i,"width"]=j_row.width;
                # dfNewSubSDS.loc[i,"codename"]=j_row.codename;
                
                # dfVeriableNonMap=dfVeriableNonMap.drop(j);
                # break;
            
    # dfNodupSDS=dfNewSubSDS.drop_duplicates(["class","form","PreText"],keep='last');
    # aaa=pd.concat([dfNewSubSDS,dfVeriableNonMap]);
    # aaa.to_csv("yyyy.csv");
    # return dfNodupSDS;
    
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
    
def createAnnotation(dfACRF,strFilePath,strFileName,strCurrentDatetime,strShowDSType,strShowVarType):
    xmlDoc=minidom.Document();
    # strCurrentDatetime=time.strftime("%Y%m%dT%H%M%S", time.localtime());
    # print("Start create Annotation!");
    # xmlInfo=xmlDoc.createElement("?xml");
    # xmlInfo.setAttribute("version","1.0");
    # xmlInfo.setAttribute("encoding","UTF-8");
    # xmlDoc.appendChild(xmlInfo);
    dfSubACRF=dfACRF[dfACRF["class"].isin(["question","unassigned"])];

    xmlXFDF=xmlDoc.createElement("xfdf");
    xmlXFDF.setAttribute("xmlns","http://ns.adobe.com/xfdf/");
    xmlXFDF.setAttribute("xml:space","preserve");

    
    xmlAnnots=xmlDoc.createElement("annots");
    strPrePage="";
    for i,i_row in dfSubACRF.iterrows():
        
        intAreaWidth=100;
        intOffset=0;
        intXaxis=i_row.x+3;
        intYaxis=i_row.y-2;
        strPage=str(i_row.pagenumber-1);
        intOrder=i_row.order;
        strAdmin=getpass.getuser();
        
        if strPage != strPrePage and str.upper(strShowDSType) =="CDASH":

            strText= i_row.CRFDS;
            strName="".join([strText,strPage,str(intOrder),"CDASHDS"]);
            strColor="#ADFF2F";
            strRect=",".join(["70","760","180","775"]);
            xmlCDASHFreeText=createFreeText(xmlDoc,strName,strRect,strText,strPage,strColor,strAdmin);
            xmlAnnots.appendChild(xmlCDASHFreeText);
                
            # elif str.upper(strShowVarType) == "BOTH" or str.upper(strShowDSType) =="SDTM":
                
                # strText= i_row.Domain;
                # strName="".join([strText,strPage,str(intOrder),"SDTMDS"]);
                # strColor="#F4D03F";
                # strRect=",".join(["180","760","360","775"]);
                # xmlCDASHFreeText=createFreeText(xmlDoc,strName,strRect,strText,strPage,strColor,strAdmin);
                # xmlAnnots.appendChild(xmlCDASHFreeText);

        if str.upper(strShowVarType) == "CDASH" or str.upper(strShowVarType) == "BOTH" :

            strRect=",".join([str(intXaxis+i_row.width),str(intYaxis),str(intXaxis+i_row.width+intAreaWidth),str(intYaxis+i_row.height)]);
            strText=i_row.CDASH;
            strColor="#ADFF2F";
            
            strName="".join([strText,strPage,str(intOrder),"CDASH"]);
            xmlCDASHFreeText=createFreeText(xmlDoc,strName,strRect,strText,strPage,strColor,strAdmin);
            xmlAnnots.appendChild(xmlCDASHFreeText);
            intOffset=intAreaWidth+2;
            
        if pd.notnull(i_row.SDTM) and i_row.SDTM !="" and (str.upper(strShowVarType) == "SDTM" or str.upper(strShowVarType) == "BOTH"):
            # print(i_row.SDTM);
            # print(i_row.CDASH);

            strRect=",".join([str(intXaxis+i_row.width+intOffset),str(intYaxis),str(intXaxis+i_row.width+50+intAreaWidth+intOffset),str(intYaxis+i_row.height)]);
            strText=i_row.SDTM;
            strName="".join([strText,strPage,str(intOrder),"SDTM"]);
            strColor="#F4D03F";
            xmlSDTMFreeText=createFreeText(xmlDoc,strName,strRect,strText,strPage,strColor,strAdmin);
            xmlAnnots.appendChild(xmlSDTMFreeText);
        
        strPrePage=strPage;
        
    xmlXFDF.appendChild(xmlAnnots);
    
    xmlF=xmlDoc.createElement("f");
    xmlF.setAttribute("href",strFileName);
    xmlXFDF.appendChild(xmlF);
    xmlDoc.appendChild(xmlXFDF);
    # print(xmlDoc.toprettyxml());
    # strXmlDoc="<?xml version='1.0' encoding='UTF-8'?>"+xmlXFDF.toprettyxml(encoding="utf-8");
    strFilename=strFileName.split(".")[0];
    # f = open(strSysPath+"/CRF/"+strFilename+strCurrentDatetime+".xfdf", "wb+");
    # f = open(strSysPath+"/CRF/"+strFilename+strCurrentDatetime+".xfdf", "wb+");
    f = open(strFilePath+"/"+strFilename+strCurrentDatetime+".xfdf", "wb+");
    
    # xmlDoc.writexml(f,encoding="utf-8");
    # print(strXmlDoc);
    # f.write(xmlDoc.toprettyxml(encoding="utf-8"));
    f.write(xmlDoc.toxml(encoding="utf-8"));
    f.close();
    # print("End create Annotation!");

def createFreeText(xmlDoc,strName,strRect,strText,strPage,strColor,strAdmin):
    strTime=time.strftime("%Y%m%d%H%M%S", time.localtime());
    xmlFreeText=xmlDoc.createElement("freetext");
    # xmlFreeText.setAttribute("color","#F4D03F");
    xmlFreeText.setAttribute("color",strColor);
    xmlFreeText.setAttribute("creationdate","D:"+strTime+"+08'00'");
    # xmlFreeText.setAttribute("flags","F4D03F");
    xmlFreeText.setAttribute("flags","000000");
    xmlFreeText.setAttribute("date","D:"+strTime+"+08'00'");
    xmlFreeText.setAttribute("name",strName);
    xmlFreeText.setAttribute("page",strPage);
    xmlFreeText.setAttribute("rect",strRect);
    xmlFreeText.setAttribute("subject","TextBox");
    xmlFreeText.setAttribute("title",strAdmin);
    
    xmlRichText=xmlDoc.createElement("contents-richtext");
    xmlBody=xmlDoc.createElement("body");
    xmlBody.setAttribute("xmlns","http://www.w3.org/1999/xhtml");
    xmlBody.setAttribute("xmlns:xfa","http://www.xfa.org/schema/xfa-data/1.0/");
    xmlBody.setAttribute("xfa:APIVersion","Acrobat:9.4.0");
    xmlBody.setAttribute("xfa:spec","2.0.2");
    xmlBody.setAttribute("style","text-align:left;color:#1C86EE;");
    xmlP=xmlDoc.createElement("p");
    xmlP.setAttribute("dir","ltr");
    
    xmlSpan=xmlDoc.createElement("span");
    xmlSpan.setAttribute("style","font-family:Helvetica");
    
    xmlSpan.appendChild(xmlDoc.createTextNode(strText));
    xmlP.appendChild(xmlSpan);
    xmlBody.appendChild(xmlP);

    xmlRichText.appendChild(xmlBody);
    
    xmlAppearance=xmlDoc.createElement("defaultappearance");
    xmlAppearance.appendChild(xmlDoc.createTextNode("0 G 205 26 28 rg 0 Tc 0 Tw 100 Tz 0 TL 0 Ts 0 Tr /Helv 12 Tf"));
    xmlStyle=xmlDoc.createElement("defaultstyle");
    xmlStyle.appendChild(xmlDoc.createTextNode("font: Helvetica,sans-serif 12.0pt; text-align:left; color:#21618C"));
    
    xmlFreeText.appendChild(xmlRichText);
    xmlFreeText.appendChild(xmlAppearance);
    xmlFreeText.appendChild(xmlStyle);
    return xmlFreeText;
    

if __name__ == '__main__':

    strSysPath=os.path.dirname(os.path.abspath(sys.argv[0]));
    multiprocessing.freeze_support();
    # print(strSysPath);
    # input("Enter:");
    fileProfile = open(strSysPath+"/init.json",'r',encoding='utf-8');
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
        pool = multiprocessing.Pool();

        dfSDS=pool.apply_async(func=parseSDS, args=(strSDSPath,strSDSName,)).get();
        dfSdsDic=pool.apply_async(func=parseSdsDic, args= (strSDSPath,strSDSName ,)).get();
        dfIgCdash=pool.apply_async(func=parseIgCdash, args=(strSysPath,"IGCDASH.json",)).get();
        dfSanofiCdash=pool.apply_async(func=parseSanofiCdash, args=(strSysPath,"SANOFICDASH.json",)).get();
        dfIgSdtm=pool.apply_async(func=parseIgSdtm, args=(strSysPath,"IGSDTM.json",)).get();
        dfSanofiSdtm=pool.apply_async(func=parseSanofiSdtm, args=(strSysPath,"SANOFISDTM.json",)).get();
        dfSanofiSdtmDom=pool.apply_async(func=parseSanofiSdtmDom, args=(strSysPath,"SANOFISDTMDOM.json",)).get();
        
        dfCRF=pool.apply_async(func=parseCRF, args=(strCRFPath,strCRFName ,)).get();
        pool.close();
        pool.join();

        dfACRF=mapSDS(dfCRF,dfSDS,dfSdsDic,dfIgCdash,dfSanofiCdash,dfIgSdtm,dfSanofiSdtm,dfSanofiSdtmDom);
        if len(strShowDSType) ==0 :
            strShowDSType="CDASH";
            
        if len(strShowVarType) ==0 :
            strShowVarType="BOTH";
        # createAnnotation(dfACRF,strCRFName,strShowVarType);
        # createAnnotation(dfACRF,strCRFPath,strCRFName,strCurrentDatetime,strShowDSType,strShowVarType);
        
        strFilename=strCRFName.split(".")[0];
        dictMeta={"NAME":strAdmin,"VERSION":"1.0.0","CREATED":strCurrentDatetime};
        dictMeta["MCRF"]=dfACRF.to_dict(orient="records");
       
        with open(strCRFPath+"/"+strFilename+strCurrentDatetime+".mcrf",'w') as f:
            # print(dfACRF.to_json(orient='records'));
            f.write(json.dumps(dictMeta));
        dfACRF.to_csv(strCRFPath+"/"+strFilename+strCurrentDatetime+".csv");
        # dictMeta.to_json(strCRFPath+"/"+strFilename+strCurrentDatetime+".json",orient='records');
        # dfACRF.to_json(strSysPath+"/data/"+strFilename+strCurrentDatetime+".json",orient='records');
        time2 = time.time();
        print(strCRFPath+"/"+strFilename+strCurrentDatetime+".mcrf");

        # print("Expended:",time2-time1);