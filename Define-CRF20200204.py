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
            # if count >5:
                # break;
                
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
    # print("SDS start Loading");
    # dfSheet=pd.read_excel(strSysPath+"/SDS/"+strFileName,["Forms","Fields"]);
    dfSheet=pd.read_excel(strFilePath+"/"+strFileName,["Forms","Fields"]);
    dfForm=dfSheet['Forms'];
    dfFields=dfSheet['Fields'];
    
    # dfForm=pd.read_excel(strFileName,sheet_name='Forms');
    # dfFields=pd.read_excel(strFileName,sheet_name='Fields');
    # dfForm,dfFields=pd.read_excel(strFileName,sheet_name=['Forms','Fields']);
    # dfFields=pd.read_excel(strFileName,sheet_name='Fields');
    
    dfSubForm=dfForm.loc[:,["OID","DraftFormName"]];
    dfSubFields=dfFields.loc[:,["FormOID","DraftFieldName","PreText","Ordinal","DataDictionaryName"]];

    dfSubFormNotNull=dfSubForm[dfSubForm["OID"].notnull()];
    dfSubFieldsNotNull=dfSubFields[dfSubFields["FormOID"].notnull()];

    # dfSubFieldsRename=dfSubFieldsNotNull.rename(columns={"FormOID":"OID","DraftFieldName":"CDASH"});
    dfSubFieldsRename=dfSubFieldsNotNull.rename(columns={"DraftFieldName":"CDASH","DataDictionaryName":"codename"});

    # dfSubFieldsRename["FormOID"].apply(lambda x: x.split("_")[0]);
    dfSubFieldsRename["DraftDomain"]=dfSubFieldsRename["FormOID"].apply(lambda x: x.split("_")[0]);
    
    dfSub=pd.merge(dfSubFormNotNull,dfSubFieldsRename,left_on="OID",right_on="FormOID",how="right");
    dfSubFormRename=dfSub.rename(columns={"DraftFormName":"form"});
    dfDraftFormNameOnly=dfSubFormRename.loc[:,["codename","form","CDASH"]].copy();

    dfSubFormRename["PreText"]=dfSubFormRename["PreText"].replace(["<i>","</i>","<i/>","<b>","</b>","<br>","</br>","<br/>","\n"],"",regex=True);
    # print("SDS finished Loading");
    return dfSubFormRename;

def parseDictionary(strFilePath,strFileName):
    # print("Dictionary finished Loading");
    # ioFileAddress = pd.io.excel.ExcelFile(strFileName);
    # dfDictionary=pd.read_excel(strFilePath+"/SDS/"+strFileName,sheet_name='DataDictionaryEntries');
    dfDictionary=pd.read_excel(strFilePath+"/"+strFileName,sheet_name='DataDictionaryEntries');
    dfSubDictionary=dfDictionary.loc[:,["DataDictionaryName","UserDataString","CodedData"]];

    dfSubDictionaryNotNull=dfSubDictionary[dfSubDictionary["DataDictionaryName"].notnull()];

    dfSubDictionaryRename=dfSubDictionaryNotNull.rename(columns={"UserDataString":"PreText","DataDictionaryName":"codename","CodedData":"codelist"});
    
    return dfSubDictionaryRename;
    # print("Dictionary finished Loading");

def parseCDASH(strFilePath,strFileName):
    # print("CDASH finished Loading");
    # ioFileAddress = pd.io.excel.ExcelFile(strFileName);
    dfCDASH=pd.read_excel(strFilePath+"/Lib/"+strFileName,sheet_name='CDASHIG_Metadata_Table');

    dfSubCDASH=dfCDASH.loc[:,["Question Text","CDASHIG Variable","SDTMIG Target"]];

    dfSubCDASHNotNull=dfSubCDASH[dfSubCDASH["CDASHIG Variable"].notnull()];

    dfSubCDASHRename=dfSubCDASHNotNull.rename(columns={"Question Text":"Question","SDTMIG Target":"SDTMIG","CDASHIG Variable":"CDASHIG"});
    dfSubCDASHAdvise=dfSubCDASHRename.apply(adviseCDASH,axis=1);
    # dfSubCDASHRename.to_csv("bbb.csv");
    # dfSubCDASHAdvise.to_csv("ccc.csv");
    dfSubCDASHNodupe=dfSubCDASHAdvise.drop_duplicates(keep="first");
    # print("CDASH finished Loading");
    return dfSubCDASHNodupe;

def adviseCDASH(x):
    if isinstance(x.SDTMIG,str):
        x["Domain"]=x.SDTMIG[0:2];
        # pattern  = re.compile(r'[A-Za-z]+(?=\.)');
        # match=pattern.search(x.SDTMIG);
        # if match:
            # x.Domain=match.group();
    return x;
    
def parseSanofiMapping(strFilePath,strFileName):
    # print("SANOFI Mapping standard start Loading");
    dfField=pd.read_excel(strFilePath+"/lib/"+strFileName,sheet_name='Fields');
    dfSubField=dfField.loc[:,["CDASHSANOFI","SDTMSANOFI","DOMAINSANOFI"]];
    dfSubFielNotNull=dfSubField[dfSubField["SDTMSANOFI"].notnull()];
    dfSubFielNodupe=dfSubFielNotNull.drop_duplicates(keep="first");
    # print("SANOFI Mapping standard end Loading");
    return dfSubFielNodupe;

def parseSanofiSDTM(strFilePath,strFileName):
    # print("SANOFI standard start Loading");
    dfField=pd.read_excel(strFilePath+"/lib/"+strFileName,sheet_name='Variable Metadata');
    dfSubField=dfField.loc[:,["Variable Name","Dataset Name"]];
    dfSubFieldRename=dfSubField.rename(columns={"Variable Name":"GSDTMSANOFI","Dataset Name":"GDOMAINSANOFI"});
    
    dfSubFielNotNull=dfSubFieldRename[dfSubFieldRename["GSDTMSANOFI"].notnull()];
    # dfSubFielAgg=dfSubFielNotNull.drop_duplicates(keep="first");
    dfSubFielAgg=dfSubFielNotNull.groupby("GSDTMSANOFI",as_index=False).agg(';'.join);
    # print(dfSubFielAgg);
    # grouped.to_csv("aaaad.csv");
    # grp=[];

    return dfSubFielAgg;
    
def mapSDS(dfCRF,dfSDS,dfDictionary,dfCDASH,dfSanofiSDTM,dfSanofiMapping):
    dfVariable=dfCRF[dfCRF["class"]=="question"].copy();
    # dfValue=dfCRF[dfCRF["class"]=="codedata"].copy();
    dfVariable["CDASHTYPE"]="SANOFI";
   
    dfVeriableMerge=pd.merge(dfVariable,dfSDS,left_on=["form","pretext"],right_on=["form","PreText"],how="left",).copy();

    dfSubDictionaryForm=pd.merge(dfDictionary,dfSDS.drop("PreText",axis=1),left_on="codename",right_on="codename",how="left").copy();
    dfSubDictionaryFormAgg=dfDictionary[["codename","codelist"]].groupby("codename",as_index=False).agg(";".join);
    # dfSubDictionaryFormAgg.to_csv("uuuuuuuuu.csv");
    dfSubSDS=dfSDS[~(dfSDS.PreText.isin(dfVeriableMerge.pretext) & dfSDS.form.isin(dfVeriableMerge.form))];
    # print(dfValue.columns);
    # print(dfSubDictionaryForm.columns);
    # dfValueMerge=pd.merge(dfValue,dfSubDictionaryForm,left_on=["form","pretext"],right_on=["form","PreText"],how="left").copy();
    # dfSubDictionary=dfSubDictionaryForm[~(dfSubDictionaryForm.PreText.isin(dfValueMerge.pretext) & dfSubDictionaryForm.form.isin(dfValueMerge.form))];
    
    # dfValueMergeNodup=dfValueMerge.drop_duplicates(keep="first");

    dfVeriableMap=dfVeriableMerge[dfVeriableMerge["CDASH"].notnull()==True].copy();
    
    dfNodupVeriableMap=dfVeriableMap.drop_duplicates(keep="first");

    dfVeriableNonMap=dfVeriableMerge[dfVeriableMerge["CDASH"].isnull()==True].copy();

    # dfValueMap=dfValueMergeNodup[dfValueMergeNodup["codename"].notnull()==True];
    # dfNodupValueMap=dfValueMap.drop_duplicates(keep="first");

    # dfValueNonMap=dfValueMergeNodup[dfValueMergeNodup["codename"].isnull()==True];

    dfNodupVeriable=mergeQuestion(dfVeriableNonMap,dfSubSDS).copy();
    # dfNodupValue=mergeCodedata(dfValueNonMap,dfSubDictionary).copy();

    # dfAllSet=pd.concat([dfNodupVeriableMap,dfNodupVeriable,dfNodupValueMap,dfNodupValue],ignore_index=True,sort=False);
    dfAllSet=pd.concat([dfNodupVeriableMap,dfNodupVeriable],ignore_index=True,sort=False);
    dfAllSet["MCDASH"]= dfAllSet["CDASH"].str.replace(r'[\(\)\d]+', '');
    
    dfMapCDASH=pd.merge(dfAllSet,dfCDASH,left_on="MCDASH",right_on="CDASHIG",how="left");
    # dfMapCDASH["CDASHTYPE"]=dfMapCDASH.apply(lambda x: "CDASHIG" if pd.notnull(x.CDASHIG) , axis=1);
    dfMapCDASH.loc[pd.notnull(dfMapCDASH.CDASHIG),"CDASHTYPE"]="IG";
    # print(dfMapCDASH.columns);
    dfMapCDASH["SDTM"]=dfMapCDASH.apply(lambda x: x.SDTMIG if pd.notnull(x.SDTMIG) else "" , axis=1);
    dfMapCDASH["SDTMTYPE"]=dfMapCDASH.apply(lambda x: "IG" if x.SDTM !="" else "" , axis=1);
    # dfSanofiSDTM.to_csv("rrrr.csv");
    dfMapSANOFI=pd.merge(dfMapCDASH,dfSanofiSDTM,left_on="CDASH",right_on="GSDTMSANOFI",how="left",copy=True);
    # dfMapSANOFI.to_csv("sdfd.csv");
    # dfMapSANOFI=pd.merge(dfMapCDASH,dfSanofiMapping,left_on="CDASH",right_on="CDASHSANOFI",how="left");
    dfMapSANOFICDASH=pd.merge(dfMapSANOFI,dfSanofiMapping,left_on="CDASH",right_on="CDASHSANOFI",how="left");
    # dfMapSANOFI["SDTMTYPE"]=dfMapSANOFI.apply(lambda x: "SANOFI" if pd.notnull(x.SDTMSANOFI) and x.SDTM=="" else x.SDTMTYPE , axis=1);
    # dfMapSANOFI["SDTM"]=dfMapSANOFI.apply(lambda x: x.SDTMSANOFI if pd.notnull(x.SDTMSANOFI) and x.SDTM=="" else x.SDTM , axis=1);
    dfMapSANOFICode=pd.merge(dfMapSANOFICDASH,dfSubDictionaryFormAgg,on="codename",how="left");
    
    dfMapSANOFIType=dfMapSANOFICode.apply(setSDTMType, axis=1);
    dfMapSANOFIFilter=dfMapSANOFIType[["pagenumber","form","OID","Domain","order","PreText","Question","CDASH","CDASHTYPE","SDTM","SDTMTYPE","codename","codelist","class","x","y","height","width"]];
    dfAllSetSort=dfMapSANOFIFilter.sort_values(by=["pagenumber","y","order"],ascending=[True,False,True]);
    dfNodupAllSetSort=dfAllSetSort[dfAllSetSort["pagenumber"].notnull()].drop_duplicates(keep="first");
    return dfNodupAllSetSort;
    
def setSDTMType(s):
    if pd.notnull(s.GSDTMSANOFI) and s.SDTM=="":
        s.SDTMTYPE="SANOFI";
        s.SDTM=s.GSDTMSANOFI;
        s.Domain=s.GDOMAINSANOFI;
    elif pd.notnull(s.SDTMSANOFI) and s.SDTM=="":
        s.SDTMTYPE="SANOFI";
        s.SDTM=s.SDTMSANOFI;
        s.Domain=s.DOMAINSANOFI;
    return s;
    
def mergeQuestion(dfVeriableNonMap,dfSubSDS):

    dfNewSubSDS=dfSubSDS.copy();
    dfTargetSDS=dfSubSDS[dfSubSDS.form.isin(dfVeriableNonMap.form)];
    # dfNewSubSDS.to_csv("b.csv");
    # dfVeriableNonMap.to_csv("a.csv");
    for i ,i_row in dfTargetSDS.iterrows():
        seriousTargetVeriable=dfVeriableNonMap[i_row.form==dfVeriableNonMap.form];
        
        dfNewSubSDS.loc[i,"order"]=0;
        dfNewSubSDS.loc[i,"x"]=-100;
        dfNewSubSDS.loc[i,"y"]=760;
        dfNewSubSDS.loc[i,"class"]="unassigned";
        dfNewSubSDS.loc[i,"height"]=15;
        dfNewSubSDS.loc[i,"width"]=100;
        for j,j_row in seriousTargetVeriable.iterrows():

            dfNewSubSDS.loc[i,"pagenumber"]=j_row.pagenumber;
            dfNewSubSDS.loc[i,"CDASHTYPE"]=j_row.CDASHTYPE;

            if (i_row.PreText in j_row.pretext) or (j_row.pretext in i_row.PreText) :
                dfNewSubSDS.loc[i,"order"]=j_row.order;
                dfNewSubSDS.loc[i,"class"]=j_row["class"];
                dfNewSubSDS.loc[i,"x"]=j_row.x;
                dfNewSubSDS.loc[i,"y"]=j_row.y;
                dfNewSubSDS.loc[i,"height"]=j_row.height;
                dfNewSubSDS.loc[i,"width"]=j_row.width;
                dfNewSubSDS.loc[i,"codename"]=j_row.codename;
                
                dfVeriableNonMap=dfVeriableNonMap.drop(j);
                break;
            
    dfNodupSDS=dfNewSubSDS.drop_duplicates(["form","PreText"],keep='last');
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

            strText= i_row.OID;
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
        dfCDASH=pool.apply_async(func=parseCDASH, args=(strSysPath,"CDASHIGv2.0_MetadataTable.xlsx",)).get();
        dfDictionary=pool.apply_async(func=parseDictionary, args= (strSDSPath,strSDSName ,)).get();
        dfCRF=pool.apply_async(func=parseCRF, args=(strCRFPath,strCRFName ,)).get();
        dfSanofiMapping=pool.apply_async(func=parseSanofiMapping, args=(strSysPath,"SanofiStandard.xlsx",)).get();
        dfSanofiSDTM=pool.apply_async(func=parseSanofiSDTM, args=(strSysPath,"Sanofi_SDTM_metadata_V1.9_SDTMIGV3.2.xlsm",)).get();
        
        pool.close();
        pool.join();
        # dfCRF["SDTM"]="SDTMTEST";
        dfACRF=mapSDS(dfCRF,dfSDS,dfDictionary,dfCDASH,dfSanofiSDTM,dfSanofiMapping);
        # dfSanofiSDTM.to_csv("bbbbbb.csv");
        if len(strShowDSType) ==0 :
            strShowDSType="CDASH";
            
        if len(strShowVarType) ==0 :
            strShowVarType="BOTH";
        # createAnnotation(dfACRF,strCRFName,strShowVarType);
        # createAnnotation(dfACRF,strCRFPath,strCRFName,strCurrentDatetime,strShowDSType,strShowVarType);
        strFilename=strCRFName.split(".")[0];
        dfACRF.to_csv(strCRFPath+"/"+strFilename+strCurrentDatetime+".csv");
        dfACRF.to_json(strCRFPath+"/"+strFilename+strCurrentDatetime+".json",orient='records');
        # dfACRF.to_csv(strSysPath+"/data/"+strFilename+strCurrentDatetime+".csv");
        dfACRF.to_json(strSysPath+"/data/"+strFilename+strCurrentDatetime+".json",orient='records');
        time2 = time.time();
        print(strCRFPath+"/"+strFilename+strCurrentDatetime+".json");
        # input("Enter:");
        # print("Expended:",time2-time1);