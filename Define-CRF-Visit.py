import sys
import time
# import xml.dom.minidom as minidom;
import os.path
import pandas as pd
import getpass
import json
import re


def parseSDS(strFilePath, strFileName):

    dictMetaDataSds = pd.read_excel(strFilePath + "\\" + strFileName, None)
    dfMatrix = pd.DataFrame()
    for key in dictMetaDataSds:
        dfSheet = dictMetaDataSds[key].rename(columns=lambda x: x.strip())
        strKey = key.upper()
        if strKey == "CRFDRAFT":
            dfDraft = dfSheet
        elif strKey == "FORMS":
            dfForm = dfSheet
        elif strKey == "FIELDS":
            dfFields = dfSheet
        elif strKey == "FOLDERS":
            dfFolders = dfSheet

        if re.match("Matrix[0-9]+", key, re.I):
            strColunmName = dfSheet.columns[0]
            dfSheetRename = dfSheet.rename(columns={strColunmName: "CRFDS"})
            dfSheetNone = dfSheetRename[dfSheetRename.drop(
                columns=["CRFDS"]).any(axis='columns')]
            # print(dfSheetRename.notnull().any(axis='columns'))
            if dfMatrix.empty:
                dfMatrix = dfSheetNone
            else:
                dfMatrix = pd.concat([dfMatrix, dfSheetNone])
            # print(dfMatrix.head())
    dfMatrixSort = dfMatrix.sort_values(by=["CRFDS"], ascending=[True])

    dfFoldersRename = dfFolders.loc[:, [
        "OID", "FolderName", "Ordinal", "Targetdays", "OverDueDays"
    ]].rename(
        columns={
            "OID": "CRFVISID",
            "FolderName": "CRFVISIT",
            "Ordinal": "CRFVISOD",
            "Targetdays": "CRFVISDY",
            "OverDueDays": "CRFVISDU"
        })

    dfMatrixAgg = dfMatrixSort.groupby("CRFDS").apply(lambda x: x.any()).drop(
        columns=["CRFDS", "Subject"])

    dfMatrixAgg["CRFVISID"] = dfMatrixAgg.apply(
        lambda x: x[x].index.str.cat(sep=','), axis=1)

    dfMatrixKeep = dfMatrixAgg["CRFVISID"].reset_index()

    dfMatrixUnstack = dfMatrixKeep.drop("CRFVISID", axis=1).join(
        dfMatrixKeep["CRFVISID"].str.split(",",
                                           expand=True).stack().reset_index(
                                               level=1,
                                               drop=True).rename("CRFVISID"))
    dfMatrixUnstack.to_csv("uuuu.csv")
    dfMatrixMerge = pd.merge(dfMatrixUnstack,
                             dfFoldersRename,
                             on="CRFVISID",
                             how="inner")
    dfMatrixMerge.to_csv("ffff.csv")

    # dfScheduleAgg = dfMatrixMerge[dfMatrixMerge["CRFVISDY"].notnull()].fillna(0).groupby(
    #     ["CRFDS"], sort=False).agg(list)
    # dfUnScheduleAgg = dfMatrixMerge[dfMatrixMerge["CRFVISDY"].isnull()].drop(
    #     columns=["CRFVISDY", "CRFVISDU"]).groupby(["CRFDS"], sort=False).agg(list)
    dfFinal = dfMatrixMerge.fillna(-1).groupby(["CRFDS"], sort=False).agg(list)

    # strStudyName=dfDraft.loc[0,"ProjectName"];
    # strVersion=dfDraft.loc[0,"DraftName"];
    # print(strStudyName);
    # print(strVersion);
    # dfFinal = dfScheduleAgg.append(dfUnScheduleAgg,sort=False)
    # dfFinal=dfMatrixMerge.groupby(["CRFDS"]).agg(list)
    dfFinal.to_csv("ccc.csv")
    # dfSubfiledSdtmDomSplit.to_csv("bbb.csv");


if __name__ == '__main__':

    strSysPath = os.path.dirname(os.path.abspath(sys.argv[0]))

    fileProfile = open(strSysPath + "\\init.json", 'r', encoding='utf-8')

    strCurrentDatetime = time.strftime("%Y%m%dT%H%M%S", time.localtime())

    jsonInfo = json.load(fileProfile)
    strCRFFullName = jsonInfo["CRF"]
    strSDSFullName = jsonInfo["SDS"]
    strShowVarType = jsonInfo["ShowVarType"]
    strShowDSType = jsonInfo["ShowDSType"]

    print("ControlVersion" in jsonInfo.keys())

    strAdmin = getpass.getuser()

    if len(strCRFFullName) > 0:
        (strCRFPath, strCRFName) = os.path.split(strCRFFullName)

    if len(strSDSFullName) > 0:
        (strSDSPath, strSDSName) = os.path.split(strSDSFullName[0])

    if len(strCRFPath) == 0:
        strCRFPath = "."

    if len(strSDSPath) == 0:
        strSDSPath = "."

    if fileProfile:
        time1 = time.time()
        # pool = multiprocessing.Pool();

        dfSDS = parseSDS(strSDSPath, strSDSName)

        time2 = time.time()

        print("Expended:", time2 - time1)
