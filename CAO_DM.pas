{



Programm     : IMMO DM
Modul        : CAO_DM
Stand        : 23.10.2004
Version      : 1.3.0.12
Beschreibung : allgemeines nichtvisuelles Datenmodul mit Grundfunktionen
               für das Programm

- Funktionen zum öffnen eines Mandanten (Datenbank im MySQL-Server)
- Funktionen zur Anlage einer neuen Datenbank und aller Tabellen + Formualre
- Funktionen für den Zugriff auf die SQL-Registry
  ( ReadString, WriteString etc. )
- Funktionen zum lesen und wirderherstellen von Tabellenlayouts (tDBGrid)
- Funktionen zum Stornieren von Vorgängen
- Funktion CalcLeitWaehrung rechnet beliebige Währung in Leitwährung um

History :

}

{*******************************************************************************

$Id: CAO_DM.pas,v 1.61.2.13 2005/04/05 06:56:55 jan Exp $

CVS-Log :
$Log: CAO_DM.pas,v $
Revision 1.61.2.13  2005/04/05 06:56:55  jan
*** empty log message ***

Revision 1.61.2.11  2005/03/11 23:04:33  NLH
no message

*******************************************************************************}

UNIT CAO_DM;

{$I CAO32.INC}

{ $DEFINE SMTPLOG}

INTERFACE

USES
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  ZTransact,
  ZMySqlTr,
  ZConnect,
  ZMySqlCon,
  ZQuery,
  ZMySqlQuery,
  Db,
  cao_var_const,
  dbgrids,
  JvStrHlder,
  JvComponent,
  JvVigenereCipher,
{$IFDEF AVE}
  cao_ave_ssh,
{$ENDIF}
  CaoSecurity,
  scExcelExport,
  IdBaseComponent,
  IdComponent,
  IdTCPConnection,
  IdTCPClient,
  IdFTP,
  IdMessageClient,
  IdSMTP,
  IdMessage,
  IdHTTP,
  FMTBcd,
  DBXpress,
  SqlExpr
  , IdSMTPBase
  , IdExplicitTLSClientServerBase
  , IdText
  ;

CONST
  berStammdatenLoeschen = 5;

TYPE
  ParseRec = RECORD
    Name : STRING;
    Wert : STRING;
  END;

TYPE
  tEkUpdateTyp = (ekuNoUpdate, ekuNewPreis, ekuMittel);

TYPE
  tWgrFaktorCache = RECORD
    Wgr : Integer;
    FTab : ARRAY[1..5] OF Double;
  END;

TYPE
  TDM1 = CLASS(TDataModule)
    LandDS                                         : TDataSource;
    WhrungDS                                       : TDataSource;
    WgrDS                                          : TDataSource;
    LiefArtDS                                      : TDataSource;
    ZahlArtDS                                      : TDataSource;
    VertreterDS                                    : TDataSource;
    Transact1                                      : TZMySqlTransact;
    JourTab                                        : TZMySqlQuery;
    JourTabQUELLE                                  : TIntegerField;
    JourTabREC_ID                                  : TIntegerField;
    JourTabVRENUM                                  : TStringField;
    JourTabKM_STAND                                : TIntegerField;
    JourTabRDATUM                                  : TDateField;
    JourTabNSUMME                                  : TFloatField;
    JourTabMSUMME                                  : TFloatField;
    JourTabBSUMME                                  : TFloatField;
    JourTabWAEHRUNG                                : TStringField;
    JourTabSTADIUM                                 : TIntegerField;
    JourTabKUN_NAME1                               : TStringField;
    JourTabPROJEKT                                 : TStringField;
    JourTabORGNUM                                  : TStringField;
    JourTabADDR_ID                                 : TIntegerField;
    JourTabKFZ_ID                                  : TIntegerField;
    JPosTab                                        : TZMySqlQuery;
    ZahlartTab                                     : TZMySqlQuery;
    VertreterTab                                   : TZMySqlQuery;
    LandTab                                        : TZMySqlQuery;
    WhrungTab                                      : TZMySqlQuery;
    WgrTab                                         : TZMySqlQuery;
    NummerTab                                      : TZMySqlQuery;
    NummerTabQUELLE                                : TIntegerField;
    KGRTab                                         : TZMySqlQuery;
    KGRTabLANGBEZ                                  : TStringField;
    KgrDS                                          : TDataSource;
    KGRTabSQL_STATEMENT                            : TMemoField;
    KGRTabGR                                       : TIntegerField;
    WhrungTabWAEHRUNG                              : TStringField;
    WhrungTabLANGBEZ                               : TStringField;
    WhrungTabFAKTOR                                : TFloatField;
    ZahlartTabZAHL_ID                              : TFloatField;
    ZahlartTabSKONTO_PROZ                          : TFloatField;
    ZahlartTabSKONTO_TAGE                          : TIntegerField;
    ZahlartTabLANGBEZ                              : TStringField;
    LiefArtTab                                     : TZMySqlQuery;
    NummerTabFORMAT                                : TStringField;
    NummerTabMAINKEY                               : TStringField;
    RegTab                                         : TZMySqlQuery;
    RegTabMAINKEY                                  : TStringField;
    RegTabNAME                                     : TStringField;
    RegTabVAL_CHAR                                 : TStringField;
    RegTabVAL_DATE                                 : TDateTimeField;
    RegTabVAL_INT                                  : TIntegerField;
    RegTabVAL_DOUBLE                               : TFloatField;
    RegTabVAL_BLOB                                 : TBlobField;
    RegTabVAL_BIN                                  : TBlobField;
    RegTabVAL_TYP                                  : TIntegerField;
    ArtMengeTab                                    : TZMySqlQuery;
    JPosTabREC_ID                                  : TIntegerField;
    JPosTabQUELLE                                  : TIntegerField;
    JPosTabQUELLE_SUB                              : TIntegerField;
    JPosTabJOURNAL_ID                              : TIntegerField;
    JPosTabARTIKELTYP                              : TStringField;
    JPosTabARTIKEL_ID                              : TIntegerField;
    JPosTabADDR_ID                                 : TIntegerField;
    JPosTabATRNUM                                  : TIntegerField;
    JPosTabVRENUM                                  : TStringField;
    JPosTabVLSNUM                                  : TStringField;
    JPosTabPOSITION                                : TIntegerField;
    JPosTabMATCHCODE                               : TStringField;
    JPosTabARTNUM                                  : TStringField;
    JPosTabBARCODE                                 : TStringField;
    JPosTabMENGE                                   : TFloatField;
    JPosTabLAENGE                                  : TStringField;
    JPosTabGROESSE                                 : TStringField;
    JPosTabDIMENSION                               : TStringField;
    JPosTabGEWICHT                                 : TFloatField;
    JPosTabME_EINHEIT                              : TStringField;
    JPosTabPR_EINHEIT                              : TFloatField;
    JPosTabEPREIS                                  : TFloatField;
    JPosTabE_RGEWINN                               : TFloatField;
    JPosTabRABATT                                  : TFloatField;
    JPosTabSTEUER_CODE                             : TIntegerField;
    JPosTabALTTEIL_PROZ                            : TFloatField;
    JPosTabALTTEIL_STCODE                          : TIntegerField;
    JPosTabGEGENKTO                                : TIntegerField;
    JPosTabBEZEICHNUNG                             : TMemoField;
    JourTabZAHLART                                 : TIntegerField;
    JourTabSOLL_STAGE                              : TIntegerField;
    JourTabSOLL_SKONTO                             : TFloatField;
    JourTabIST_ANZAHLUNG                           : TFloatField;
    JourTabIST_SKONTO                              : TFloatField;
    JourTabIST_ZAHLDAT                             : TDateField;
    JourTabIST_BETRAG                              : TFloatField;
    JourTabKONTOAUSZUG                             : TIntegerField;
    JourTabBANK_ID                                 : TIntegerField;
    JourTabUW_NUM                                  : TIntegerField;
    CpySrcKopfTab                                  : TZMySqlQuery;
    CpyDstKopfTab                                  : TZMySqlQuery;
    CpySrcPosTab                                   : TZMySqlQuery;
    CpyDstPosTab                                   : TZMySqlQuery;
    CpySrcKopfTabQUELLE                            : TIntegerField;
    CpySrcKopfTabREC_ID                            : TIntegerField;
    CpySrcKopfTabQUELLE_SUB                        : TIntegerField;
    CpySrcKopfTabADDR_ID                           : TIntegerField;
    CpySrcKopfTabATRNUM                            : TIntegerField;
    CpySrcKopfTabVRENUM                            : TStringField;
    CpySrcKopfTabVLSNUM                            : TStringField;
    CpySrcKopfTabFOLGENR                           : TIntegerField;
    CpySrcKopfTabKM_STAND                          : TIntegerField;
    CpySrcKopfTabKFZ_ID                            : TIntegerField;
    CpySrcKopfTabVERTRETER_ID                      : TIntegerField;
    CpySrcKopfTabGLOBRABATT                        : TFloatField;
    CpySrcKopfTabADATUM                            : TDateField;
    CpySrcKopfTabRDATUM                            : TDateField;
    CpySrcKopfTabLDATUM                            : TDateField;
    CpySrcKopfTabTermin                            : TDateField;
    CpySrcKopfTabPR_EBENE                          : TIntegerField;
    CpySrcKopfTabLIEFART                           : TIntegerField;
    CpySrcKopfTabZAHLART                           : TIntegerField;
    CpySrcKopfTabKOST_NETTO                        : TFloatField;
    CpySrcKopfTabWERT_NETTO                        : TFloatField;
    CpySrcKopfTabLOHN                              : TFloatField;
    CpySrcKopfTabWARE                              : TFloatField;
    CpySrcKopfTabTKOST                             : TFloatField;
    CpySrcKopfTabMWST_0                            : TFloatField;
    CpySrcKopfTabMWST_1                            : TFloatField;
    CpySrcKopfTabMWST_2                            : TFloatField;
    CpySrcKopfTabMWST_3                            : TFloatField;
    CpySrcKopfTabNSUMME                            : TFloatField;
    CpySrcKopfTabMSUMME_0                          : TFloatField;
    CpySrcKopfTabMSUMME_1                          : TFloatField;
    CpySrcKopfTabMSUMME_2                          : TFloatField;
    CpySrcKopfTabMSUMME_3                          : TFloatField;
    CpySrcKopfTabMSUMME                            : TFloatField;
    CpySrcKopfTabBSUMME                            : TFloatField;
    CpySrcKopfTabATSUMME                           : TFloatField;
    CpySrcKopfTabATMSUMME                          : TFloatField;
    CpySrcKopfTabWAEHRUNG                          : TStringField;
    CpySrcKopfTabGEGENKONTO                        : TIntegerField;
    CpySrcKopfTabSOLL_STAGE                        : TIntegerField;
    CpySrcKopfTabSOLL_SKONTO                       : TFloatField;
    CpySrcKopfTabSOLL_NTAGE                        : TIntegerField;
    CpySrcKopfTabSOLL_RATEN                        : TIntegerField;
    CpySrcKopfTabSOLL_RATBETR                      : TFloatField;
    CpySrcKopfTabSOLL_RATINTERVALL                 : TIntegerField;
    CpySrcKopfTabIST_ANZAHLUNG                     : TFloatField;
    CpySrcKopfTabIST_SKONTO                        : TFloatField;
    CpySrcKopfTabIST_ZAHLDAT                       : TDateField;
    CpySrcKopfTabIST_BETRAG                        : TFloatField;
    CpySrcKopfTabMAHNKOSTEN                        : TFloatField;
    CpySrcKopfTabKONTOAUSZUG                       : TIntegerField;
    CpySrcKopfTabBANK_ID                           : TIntegerField;
    CpySrcKopfTabSTADIUM                           : TIntegerField;
    CpySrcKopfTabERSTELLT                          : TDateField;
    CpySrcKopfTabERST_NAME                         : TStringField;
    CpySrcKopfTabKUN_ANREDE                        : TStringField;
    CpySrcKopfTabKUN_NAME1                         : TStringField;
    CpySrcKopfTabKUN_NAME2                         : TStringField;
    CpySrcKopfTabKUN_NAME3                         : TStringField;
    CpySrcKopfTabKUN_ABTEILUNG                     : TStringField;
    CpySrcKopfTabKUN_STRASSE                       : TStringField;
    CpySrcKopfTabKUN_LAND                          : TStringField;
    CpySrcKopfTabKUN_PLZ                           : TStringField;
    CpySrcKopfTabKUN_ORT                           : TStringField;
    CpySrcKopfTabUSR1                              : TStringField;
    CpySrcKopfTabUSR2                              : TStringField;
    CpySrcKopfTabPROJEKT                           : TStringField;
    CpySrcKopfTabORGNUM                            : TStringField;
    CpySrcKopfTabBEST_NAME                         : TStringField;
    CpySrcKopfTabBEST_CODE                         : TIntegerField;
    CpySrcKopfTabINFO                              : TMemoField;
    CpySrcKopfTabUW_NUM                            : TIntegerField;
    CpyDstKopfTabQUELLE                            : TIntegerField;
    CpyDstKopfTabREC_ID                            : TIntegerField;
    CpyDstKopfTabQUELLE_SUB                        : TIntegerField;
    CpyDstKopfTabADDR_ID                           : TIntegerField;
    CpyDstKopfTabATRNUM                            : TIntegerField;
    CpyDstKopfTabVRENUM                            : TStringField;
    CpyDstKopfTabVLSNUM                            : TStringField;
    CpyDstKopfTabFOLGENR                           : TIntegerField;
    CpyDstKopfTabKM_STAND                          : TIntegerField;
    CpyDstKopfTabKFZ_ID                            : TIntegerField;
    CpyDstKopfTabVERTRETER_ID                      : TIntegerField;
    CpyDstKopfTabGLOBRABATT                        : TFloatField;
    CpyDstKopfTabADATUM                            : TDateField;
    CpyDstKopfTabRDATUM                            : TDateField;
    CpyDstKopfTabLDATUM                            : TDateField;
    CpyDstKopfTabTermin                            : TDateField;
    CpyDstKopfTabPR_EBENE                          : TIntegerField;
    CpyDstKopfTabLIEFART                           : TIntegerField;
    CpyDstKopfTabZAHLART                           : TIntegerField;
    CpyDstKopfTabKOST_NETTO                        : TFloatField;
    CpyDstKopfTabWERT_NETTO                        : TFloatField;
    CpyDstKopfTabLOHN                              : TFloatField;
    CpyDstKopfTabWARE                              : TFloatField;
    CpyDstKopfTabTKOST                             : TFloatField;
    CpyDstKopfTabMWST_0                            : TFloatField;
    CpyDstKopfTabMWST_1                            : TFloatField;
    CpyDstKopfTabMWST_2                            : TFloatField;
    CpyDstKopfTabMWST_3                            : TFloatField;
    CpyDstKopfTabNSUMME                            : TFloatField;
    CpyDstKopfTabMSUMME_0                          : TFloatField;
    CpyDstKopfTabMSUMME_1                          : TFloatField;
    CpyDstKopfTabMSUMME_2                          : TFloatField;
    CpyDstKopfTabMSUMME_3                          : TFloatField;
    CpyDstKopfTabMSUMME                            : TFloatField;
    CpyDstKopfTabBSUMME                            : TFloatField;
    CpyDstKopfTabATSUMME                           : TFloatField;
    CpyDstKopfTabATMSUMME                          : TFloatField;
    CpyDstKopfTabWAEHRUNG                          : TStringField;
    CpyDstKopfTabGEGENKONTO                        : TIntegerField;
    CpyDstKopfTabSOLL_STAGE                        : TIntegerField;
    CpyDstKopfTabSOLL_SKONTO                       : TFloatField;
    CpyDstKopfTabSOLL_NTAGE                        : TIntegerField;
    CpyDstKopfTabSOLL_RATEN                        : TIntegerField;
    CpyDstKopfTabSOLL_RATBETR                      : TFloatField;
    CpyDstKopfTabSOLL_RATINTERVALL                 : TIntegerField;
    CpyDstKopfTabIST_ANZAHLUNG                     : TFloatField;
    CpyDstKopfTabIST_SKONTO                        : TFloatField;
    CpyDstKopfTabIST_ZAHLDAT                       : TDateField;
    CpyDstKopfTabIST_BETRAG                        : TFloatField;
    CpyDstKopfTabMAHNKOSTEN                        : TFloatField;
    CpyDstKopfTabKONTOAUSZUG                       : TIntegerField;
    CpyDstKopfTabBANK_ID                           : TIntegerField;
    CpyDstKopfTabSTADIUM                           : TIntegerField;
    CpyDstKopfTabERSTELLT                          : TDateField;
    CpyDstKopfTabERST_NAME                         : TStringField;
    CpyDstKopfTabKUN_ANREDE                        : TStringField;
    CpyDstKopfTabKUN_NAME1                         : TStringField;
    CpyDstKopfTabKUN_NAME2                         : TStringField;
    CpyDstKopfTabKUN_NAME3                         : TStringField;
    CpyDstKopfTabKUN_ABTEILUNG                     : TStringField;
    CpyDstKopfTabKUN_STRASSE                       : TStringField;
    CpyDstKopfTabKUN_LAND                          : TStringField;
    CpyDstKopfTabKUN_PLZ                           : TStringField;
    CpyDstKopfTabKUN_ORT                           : TStringField;
    CpyDstKopfTabUSR1                              : TStringField;
    CpyDstKopfTabUSR2                              : TStringField;
    CpyDstKopfTabPROJEKT                           : TStringField;
    CpyDstKopfTabORGNUM                            : TStringField;
    CpyDstKopfTabBEST_NAME                         : TStringField;
    CpyDstKopfTabBEST_CODE                         : TIntegerField;
    CpyDstKopfTabINFO                              : TMemoField;
    CpyDstKopfTabUW_NUM                            : TIntegerField;
    CpySrcKopfTabKUN_NUM                           : TStringField;
    CpySrcKopfTabBEST_DATUM                        : TDateField;
    CpyDstKopfTabKUN_NUM                           : TStringField;
    CpyDstKopfTabBEST_DATUM                        : TDateField;
    KasBuch                                        : TZMySqlQuery;
    KasBuchBDATUM                                  : TDateField;
    KasBuchQUELLE                                  : TIntegerField;
    KasBuchJOURNAL_ID                              : TIntegerField;
    KasBuchZU_ABGANG                               : TFloatField;
    KasBuchBTXT                                    : TMemoField;
    KasBuchBELEGNUM                                : TStringField;
    KasBuchGKONTO                                  : TIntegerField;
    KasBuchSKONTO                                  : TFloatField;
    JourTabGEGENKONTO                              : TIntegerField;
    JourTabKUN_NUM                                 : TStringField;
    JourTabKUN_ANREDE                              : TStringField;
    JourTabKUN_NAME2                               : TStringField;
    JourTabKUN_NAME3                               : TStringField;
    JourTabKUN_ABTEILUNG                           : TStringField;
    JourTabKUN_STRASSE                             : TStringField;
    JourTabKUN_LAND                                : TStringField;
    JourTabKUN_PLZ                                 : TStringField;
    JourTabKUN_ORT                                 : TStringField;
    JourTabSOLL_NTAGE                              : TIntegerField;
    JourTabSOLL_RATEN                              : TIntegerField;
    JourTabSOLL_RATBETR                            : TFloatField;
    JourTabSOLL_RATINTERVALL                       : TIntegerField;
    JourTabIST_SKONTO_BETR                         : TFloatField;
    STListTab                                      : TZMySqlQuery;
    STListTabREC_ID                                : TIntegerField;
    STListTabART_ID                                : TIntegerField;
    STListTabMENGE                                 : TFloatField;
    STListTabERSTELLT                              : TDateField;
    STListTabERST_NAME                             : TStringField;
    STListTabGEAEND                                : TDateField;
    STListTabGEAEND_NAME                           : TStringField;
    JPosTabVIEW_POS                                : TStringField;
    KunTab                                         : TZMySqlQuery;
    JourTabVLSNUM                                  : TStringField;
    JourTabLDATUM                                  : TDateField;
    JourTabLIEFART                                 : TIntegerField;
    JourTabKOST_NETTO                              : TFloatField;
    JourTabWERT_NETTO                              : TFloatField;
    ZahlartTabTEXT                                 : TMemoField;
    NummerTabNAME                                  : TStringField;
    CreateMandantStr                               : TJvStrHolder;
    JourTabQUELLE_SUB                              : TIntegerField;
    KGRTabMAINKEY                                  : TStringField;
    FirBankTab                                     : TZMySqlQuery;
    FirBankTabkurzbez                              : TStringField;
    FirBankTabinhaber                              : TStringField;
    FirBankTabblz                                  : TIntegerField;
    FirBankTabmainkey                              : TStringField;
    UpdateArtTab                                   : TZMySqlQuery;
    ReKFZTab                                       : TZMySqlQuery;
    ReKFZTabKFZ_ID                                 : TIntegerField;
    ReKFZTabADDR_ID                                : TIntegerField;
    ReKFZTabFGST_NUM                               : TStringField;
    ReKFZTabPOL_KENNZ                              : TStringField;
    ReKFZTabSCHL_ZU_2                              : TStringField;
    ReKFZTabSCHL_ZU_3                              : TStringField;
    ReKFZTabKM_STAND                               : TIntegerField;
    ReKFZTabZULASSUNG                              : TDateField;
    ReKFZTabLE_BESUCH                              : TDateField;
    ReKFZTabNAE_TUEV                               : TDateField;
    ReKFZTabNAE_AU                                 : TDateField;
    JourTabMAHNKOSTEN                              : TFloatField;
    CpySrcKopfTabLIEF_ADDR_ID                      : TIntegerField;
    CpySrcKopfTabMAHNSTUFE                         : TIntegerField;
    CpySrcKopfTabMAHNDATUM                         : TDateField;
    CpySrcKopfTabMAHNPRINT                         : TIntegerField;
    CpyDstKopfTabLIEF_ADDR_ID                      : TIntegerField;
    CpyDstKopfTabMAHNSTUFE                         : TIntegerField;
    CpyDstKopfTabMAHNDATUM                         : TDateField;
    CpyDstKopfTabMAHNPRINT                         : TIntegerField;
    LiefRabGrp                                     : TZMySqlQuery;
    LiefRabGrpRABGRP_ID                            : TStringField;
    LiefRabGrpRABGRP_TYP                           : TIntegerField;
    LiefRabGrpMIN_MENGE                            : TIntegerField;
    LiefRabGrpLIEF_RABGRP                          : TIntegerField;
    LiefRabGrpRABATT1                              : TFloatField;
    LiefRabGrpRABATT2                              : TFloatField;
    LiefRabGrpRABATT3                              : TFloatField;
    LiefRabGrpADDR_ID                              : TIntegerField;
    LiefRabGrpBESCHREIBUNG                         : TStringField;
    RabGrpDS                                       : TDataSource;
    ZBatchSql1                                     : TZBatchSql;
    FirmaTab                                       : TZMySqlQuery;
    FirmaDS                                        : TDataSource;
    JPosTabGEBUCHT                                 : TBooleanField;
    JPosTabSN_FLAG                                 : TBooleanField;
    JPosTabALTTEIL_FLAG                            : TBooleanField;
    JPosTabBEZ_FEST_FLAG                           : TBooleanField;
    JourTabFREIGABE1_FLAG                          : TBooleanField;
    CpySrcKopfTabFREIGABE1_FLAG                    : TBooleanField;
    CpyDstKopfTabFREIGABE1_FLAG                    : TBooleanField;
    JPosTabQUELLE_SRC                              : TIntegerField;
    JourTabBRUTTO_FLAG                             : TBooleanField;
    JourTabMWST_FREI_FLAG                          : TBooleanField;
    JPosTabBRUTTO_FLAG                             : TBooleanField;
    JPosTabRABATT2                                 : TFloatField;
    JPosTabRABATT3                                 : TFloatField;
    CpySrcKopfTabBRUTTO_FLAG                       : TBooleanField;
    CpySrcKopfTabMWST_FREI_FLAG                    : TBooleanField;
    CpyDstKopfTabBRUTTO_FLAG                       : TBooleanField;
    CpyDstKopfTabMWST_FREI_FLAG                    : TBooleanField;
    JourTabMWST_0                                  : TFloatField;
    JourTabMWST_1                                  : TFloatField;
    JourTabMWST_2                                  : TFloatField;
    JourTabMWST_3                                  : TFloatField;
    JourTabMSUMME_0                                : TFloatField;
    JourTabMSUMME_1                                : TFloatField;
    JourTabMSUMME_2                                : TFloatField;
    JourTabMSUMME_3                                : TFloatField;
    KunRabGrp                                      : TZMySqlQuery;
    KunRabGrpRABGRP_ID                             : TStringField;
    KunRabGrpRABGRP_TYP                            : TIntegerField;
    KunRabGrpMIN_MENGE                             : TIntegerField;
    KunRabGrpLIEF_RABGRP                           : TIntegerField;
    KunRabGrpRABATT1                               : TFloatField;
    KunRabGrpRABATT2                               : TFloatField;
    KunRabGrpRABATT3                               : TFloatField;
    KunRabGrpADDR_ID                               : TIntegerField;
    UniQuery                                       : TZMySqlQuery;
    UniQuery2                                      : TZMySqlQuery;
    ShopOrderStatusTab                             : TZMySqlQuery;
    ShopOrderStatusTabLANGBEZ                      : TStringField;
    ShopOrderStatusTabTEXT                         : TMemoField;
    ShopOrderStatusTabMAINKEY                      : TStringField;
    ShopOSDS                                       : TDataSource;
    ShopOrderStatusTabORDERSTATUS_ID               : TIntegerField;
    HerstellerTab                                  : TZMySqlQuery;
    HerstellerDS                                   : TDataSource;
    WgrTabID                                       : TIntegerField;
    WgrTabTOP_ID                                   : TIntegerField;
    WgrTabNAME                                     : TStringField;
    WgrTabBESCHREIBUNG                             : TMemoField;
    WgrTabDEF_EKTO                                 : TIntegerField;
    WgrTabDEF_AKTO                                 : TIntegerField;
    WgrTabVORGABEN                                 : TMemoField;
    SprachTab                                      : TZMySqlQuery;
    SprachDS                                       : TDataSource;
    FirmaTabANREDE                                 : TStringField;
    FirmaTabNAME1                                  : TStringField;
    FirmaTabNAME2                                  : TStringField;
    FirmaTabNAME3                                  : TStringField;
    FirmaTabSTRASSE                                : TStringField;
    FirmaTabLAND                                   : TStringField;
    FirmaTabPLZ                                    : TStringField;
    FirmaTabORT                                    : TStringField;
    FirmaTabVORWAHL                                : TStringField;
    FirmaTabTELEFON1                               : TStringField;
    FirmaTabTELEFON2                               : TStringField;
    FirmaTabMOBILFUNK                              : TStringField;
    FirmaTabFAX                                    : TStringField;
    FirmaTabEMAIL                                  : TStringField;
    FirmaTabWEBSEITE                               : TStringField;
    FirmaTabBANK1_BLZ                              : TStringField;
    FirmaTabBANK1_KONTONR                          : TStringField;
    FirmaTabBANK1_NAME                             : TStringField;
    FirmaTabBANK1_IBAN                             : TStringField;
    FirmaTabBANK1_SWIFT                            : TStringField;
    FirmaTabBANK2_BLZ                              : TStringField;
    FirmaTabBANK2_KONTONR                          : TStringField;
    FirmaTabBANK2_NAME                             : TStringField;
    FirmaTabBANK2_IBAN                             : TStringField;
    FirmaTabBANK2_SWIFT                            : TStringField;
    FirmaTabKOPFTEXT                               : TMemoField;
    FirmaTabFUSSTEXT                               : TMemoField;
    FirmaTabABSENDER                               : TStringField;
    FirmaTabSTEUERNUMMER                           : TStringField;
    FirmaTabUST_ID                                 : TStringField;
    FirmaTabIMAGE1                                 : TBlobField;
    FirmaTabIMAGE2                                 : TBlobField;
    FirmaTabIMAGE3                                 : TBlobField;
    FirmaTabUSER_AKT                               : TStringField;
    FirmaTabLEITWAEHRUNG                           : TStringField;
    FirmaTabMANDANT_NAME                           : TStringField;
    JourTabLIEF_ADDR_ID                            : TIntegerField;
    CpySrcKopfTabPROVIS_WERT                       : TFloatField;
    CpyDstKopfTabPROVIS_WERT                       : TFloatField;
    NummerTabNEXT_NUM                              : TLargeintField;
    RegTabVAL_INT2                                 : TLargeintField;
    RegTabVAL_INT3                                 : TLargeintField;
    FirBankTabKTONR                                : TLargeintField;
    ZahlartTabNETTO_TAGE                           : TLargeintField;
    ArtInfoTab                                     : TZMySqlQuery;
    ArtInfoTabREC_ID                               : TIntegerField;
    ArtInfoTabEK_PREIS                             : TFloatField;
    ArtInfoTabVK1                                  : TFloatField;
    ArtInfoTabVK2                                  : TFloatField;
    ArtInfoTabVK3                                  : TFloatField;
    ArtInfoTabVK4                                  : TFloatField;
    ArtInfoTabVK5                                  : TFloatField;
    ArtInfoTabVK1B                                 : TFloatField;
    ArtInfoTabVK2B                                 : TFloatField;
    ArtInfoTabVK3B                                 : TFloatField;
    ArtInfoTabVK4B                                 : TFloatField;
    ArtInfoTabVK5B                                 : TFloatField;
    ArtInfoTabMENGE2                               : TIntegerField;
    ArtInfoTabPREIS2                               : TFloatField;
    ArtInfoTabMENGE3                               : TIntegerField;
    ArtInfoTabPREIS3                               : TFloatField;
    ArtInfoTabMENGE4                               : TIntegerField;
    ArtInfoTabPREIS4                               : TFloatField;
    ArtInfoTabMENGE5                               : TIntegerField;
    ArtInfoTabPREIS5                               : TFloatField;
    ArtInfoTabMENGE_AKT                            : TFloatField;
    ArtInfoTabRABGRP_ID                            : TStringField;
    ArtInfoTabMENGE_BESTELLT                       : TFloatField;
    ArtInfoTabMENGE_RESERVIERT                     : TFloatField;
    ArtInfoTabPROVIS_PROZ                          : TFloatField;
    ArtInfoTabSTEUER_CODE                          : TIntegerField;
    ArtInfoTabERLOES_KTO                           : TIntegerField;
    ArtInfoTabAUFW_KTO                             : TIntegerField;
    ArtInfoTabARTNUM                               : TStringField;
    ArtInfoTabERSATZ_ARTNUM                        : TStringField;
    ArtInfoTabMATCHCODE                            : TStringField;
    ArtInfoTabWARENGRUPPE                          : TIntegerField;
    ArtInfoTabBARCODE                              : TStringField;
    ArtInfoTabARTIKELTYP                           : TStringField;
    ArtInfoTabKAS_NAME                             : TStringField;
    ArtInfoTabME_EINHEIT                           : TStringField;
    ArtInfoTabPR_EINHEIT                           : TFloatField;
    ArtInfoTabLAENGE                               : TStringField;
    ArtInfoTabGROESSE                              : TStringField;
    ArtInfoTabDIMENSION                            : TStringField;
    ArtInfoTabGEWICHT                              : TFloatField;
    ArtInfoTabKURZNAME                             : TStringField;
    ArtInfoTabLANGNAME                             : TMemoField;
    ReKunTab                                       : TZMySqlQuery;
    WgrTabSTEUER_CODE                              : TIntegerField;
    WgrTabVK1_FAKTOR                               : TFloatField;
    WgrTabVK2_FAKTOR                               : TFloatField;
    WgrTabVK3_FAKTOR                               : TFloatField;
    WgrTabVK4_FAKTOR                               : TFloatField;
    WgrTabVK5_FAKTOR                               : TFloatField;
    NummerTabMAXLEN                                : TLargeintField;
    ArtInfoTabPREIS                                : TFloatField;
    ArtInfoTabPREIS_TYP                            : TIntegerField;
    ArtInfoTabADRESS_ID                            : TIntegerField;
    ArtInfoTabVPE                                  : TIntegerField;
    db1                                            : TZMySqlDatabase;
    JourTabSHOP_ID                                 : TIntegerField;
    JourTabSHOP_ORDERID                            : TIntegerField;
    CpyDstKopfTabGEWICHT                           : TFloatField;
    CpyDstKopfTabROHGEWINN                         : TFloatField;
    CpySrcKopfTabGEWICHT                           : TFloatField;
    CpySrcKopfTabROHGEWINN                         : TFloatField;
    DBUpDTo1_10                                    : TJvStrHolder;
    Cipher                                         : TJvVigenereCipher;
    LandTabID                                      : TStringField;
    LandTabNAME                                    : TStringField;
    LandTabISO_CODE_3                              : TStringField;
    LandTabFORMAT                                  : TIntegerField;
    LandTabVORWAHL                                 : TStringField;
    LandTabWAEHRUNG                                : TStringField;
    LandTabSPRACHE                                 : TStringField;
    CaoSecurity                                    : tCaoSecurity;
    FirBankTabFIBU_KTO                             : TLargeintField;
    KasBuchMA_ID                                   : TIntegerField;
    KasBuchERSTELLT                                : TDateField;
    KasBuchERST_NAME                               : TStringField;
    ArtInfoTabBREITE                               : TStringField;
    ArtInfoTabHOEHE                                : TStringField;
    CpySrcKopfTabNSUMME_0                          : TFloatField;
    CpySrcKopfTabNSUMME_1                          : TFloatField;
    CpySrcKopfTabNSUMME_2                          : TFloatField;
    CpySrcKopfTabNSUMME_3                          : TFloatField;
    CpySrcKopfTabBSUMME_0                          : TFloatField;
    CpySrcKopfTabBSUMME_1                          : TFloatField;
    CpySrcKopfTabBSUMME_2                          : TFloatField;
    CpySrcKopfTabBSUMME_3                          : TFloatField;
    CpyDstKopfTabNSUMME_0                          : TFloatField;
    CpyDstKopfTabNSUMME_1                          : TFloatField;
    CpyDstKopfTabNSUMME_2                          : TFloatField;
    CpyDstKopfTabNSUMME_3                          : TFloatField;
    CpyDstKopfTabBSUMME_0                          : TFloatField;
    CpyDstKopfTabBSUMME_1                          : TFloatField;
    CpyDstKopfTabBSUMME_2                          : TFloatField;
    CpyDstKopfTabBSUMME_3                          : TFloatField;
    LockQuery                                      : TZMySqlQuery;
    LandTabPOST_CODE                               : TStringField;
    LandTabEU_LAND                                 : TStringField;
    JourTabGEWICHT                                 : TFloatField;
    JourTabLOHN                                    : TFloatField;
    JourTabWARE                                    : TFloatField;
    JourTabROHGEWINN                               : TFloatField;
    JourTabTKOST                                   : TFloatField;
    JourTabNSUMME_0                                : TFloatField;
    JourTabNSUMME_1                                : TFloatField;
    JourTabNSUMME_2                                : TFloatField;
    JourTabNSUMME_3                                : TFloatField;
    JourTabBSUMME_0                                : TFloatField;
    JourTabBSUMME_1                                : TFloatField;
    JourTabBSUMME_2                                : TFloatField;
    JourTabBSUMME_3                                : TFloatField;
    JourTabPROVIS_WERT                             : TFloatField;
    ArtikelBilderQuery                             : TZMySqlQuery;
    DBUpDTo1_11                                    : TJvStrHolder;
    scExcelExport1                                 : TscExcelExport;
    JPosTabGPREIS                                  : TFloatField;
    DBupDToIMMO                                    : TJvStrHolder;
    IdFTP1                                         : TIdFTP;
    MitarbeiterTab                                 : TZMySqlQuery;
    MitarbeiterTabidmitarbeiter                    : TIntegerField;
    MitarbeiterTabaktiv                            : TIntegerField;
    MitarbeiterTablogin                            : TStringField;
    MitarbeiterDS                                  : TDataSource;
    ObjGruppenTab                                  : TZMySqlQuery;
    ObjGruppenTabRecID                             : TIntegerField;
    ObjGruppenTabTOP_ID                            : TIntegerField;
    ObjGruppenTabNAME                              : TStringField;
    ObjGruppenDS                                   : TDataSource;
    DBUpDTo1_13                                    : TJvStrHolder;
    FirmaTabeMailPOP3                              : TStringField;
    FirmaTabeMailSMTP                              : TStringField;
    FirmaTabeMailBenutzer                          : TStringField;
    FirmaTabeMailPasswort                          : TStringField;
    MitarbeiterAktivTab                            : TZMySqlQuery;
    MitarbeiterAktivDS                             : TDataSource;
    MitarbeiterTabKalenderfarbe                    : TFloatField;
    MitarbeiterTabKurz                             : TStringField;
    FirmaTabeMailPOP3Anfragen                      : TStringField;
    FirmaTabeMailSMTPAnfragen                      : TStringField;
    FirmaTabeMailBenutzerAnfragen                  : TStringField;
    FirmaTabeMailPasswortAnfragen                  : TStringField;
    FirmaTabeMailPOP3AnfragenSicherung             : TStringField;
    FirmaTabeMailSMTPAnfragenSicherung             : TStringField;
    FirmaTabeMailBenutzerAnfragenSicherung         : TStringField;
    FirmaTabeMailPasswortAnfragenSicherung         : TStringField;
    emMessage                                      : TIdMessage;
    emSMTP                                         : TIdSMTP;
    IdHTTP1                                        : TIdHTTP;
    FirmaTabSMSURL                                 : TStringField;
    FirmaTabSMSBenutzerName                        : TStringField;
    FirmaTabSMSPasswort                            : TStringField;
    KunTabkundennr                                 : TIntegerField;
    KunTabalter1                                   : TStringField;
    KunTabsuchend                                  : TIntegerField;
    KunTabstatus                                   : TIntegerField;
    KunTabidkundenstatus                           : TIntegerField;
    KunTabidkundenart                              : TIntegerField;
    KunTabtitel1                                   : TStringField;
    KunTabname1                                    : TStringField;
    KunTabvorname1                                 : TStringField;
    KunTabtitel2                                   : TStringField;
    KunTabname2                                    : TStringField;
    KunTabvorname2                                 : TStringField;
    KunTabstrasse                                  : TStringField;
    KunTabhausnr                                   : TStringField;
    KunTabort                                      : TStringField;
    KunTabtel_gesch_land                           : TStringField;
    KunTabtel_gesch_vw                             : TStringField;
    KunTabtel_gesch_nr                             : TStringField;
    KunTabtel_gesch                                : TStringField;
    KunTabtel_privat_land                          : TStringField;
    KunTabtel_privat_vw                            : TStringField;
    KunTabtel_privat_nr                            : TStringField;
    KunTabtel_privat                               : TStringField;
    KunTabtel_mobil_land                           : TStringField;
    KunTabtel_mobil_vw                             : TStringField;
    KunTabtel_mobil_nr                             : TStringField;
    KunTabtel_mobil                                : TStringField;
    KunTabtel_gesch2_land                          : TStringField;
    KunTabtel_gesch2_vw                            : TStringField;
    KunTabtel_gesch2_nr                            : TStringField;
    KunTabtel_gesch2                               : TStringField;
    KunTabtel_privat2_land                         : TStringField;
    KunTabtel_privat2_vw                           : TStringField;
    KunTabtel_privat2_nr                           : TStringField;
    KunTabtel_privat2                              : TStringField;
    KunTabtel_mobil2_land                          : TStringField;
    KunTabtel_mobil2_vw                            : TStringField;
    KunTabtel_mobil2_nr                            : TStringField;
    KunTabtel_mobil2                               : TStringField;
    KunTabfax_land                                 : TStringField;
    KunTabfax_vw                                   : TStringField;
    KunTabfax_nr                                   : TStringField;
    KunTabfax                                      : TStringField;
    KunTabfax2_land                                : TStringField;
    KunTabfax2_vw                                  : TStringField;
    KunTabfax2_nr                                  : TStringField;
    KunTabfax2                                     : TStringField;
    KunTabemail                                    : TStringField;
    KunTabemail2                                   : TStringField;
    KunTabgeb1                                     : TStringField;
    KunTabalter2                                   : TIntegerField;
    KunTabgeb2                                     : TStringField;
    KunTabnationalitaet1                           : TStringField;
    KunTabnationalitaet2                           : TStringField;
    KunTabidfamilienstand                          : TIntegerField;
    KunTabkinder                                   : TIntegerField;
    KunTabkinder_anz                               : TIntegerField;
    KunTabkinder_alter                             : TIntegerField;
    KunTabkinder_alter2                            : TIntegerField;
    KunTabberuf1                                   : TStringField;
    KunTabbranche1                                 : TStringField;
    KunTabfirma1                                   : TStringField;
    KunTabberuf_selb1                              : TIntegerField;
    KunTabberuf2                                   : TStringField;
    KunTabbranche2                                 : TStringField;
    KunTabfirma2                                   : TStringField;
    KunTabberuf_selb2                              : TIntegerField;
    KunTabverdienst1                               : TStringField;
    KunTabverdienst2                               : TIntegerField;
    KunTabidobjektkategorie                        : TIntegerField;
    KunTabobj_art                                  : TStringField;
    KunTabidobjgenart                              : TIntegerField;
    KunTabobj_gen_kauf                             : TIntegerField;
    KunTabobj_gen_miete                            : TIntegerField;
    KunTabidzimmer                                 : TIntegerField;
    KunTabidzimmer2                                : TIntegerField;
    KunTabobj_plz                                  : TIntegerField;
    KunTabobj_ort                                  : TStringField;
    KunTabobj_ort1                                 : TStringField;
    KunTabobj_ort2                                 : TStringField;
    KunTabobj_ort3                                 : TStringField;
    KunTabobj_ort4                                 : TStringField;
    KunTablkr                                      : TStringField;
    KunTablkr2                                     : TStringField;
    KunTablkr3                                     : TStringField;
    KunTabobj_flaeche_min                          : TIntegerField;
    KunTabobj_flaeche_max                          : TIntegerField;
    KunTabpreis_min                                : TFloatField;
    KunTabpreis_max                                : TFloatField;
    KunTabkommentar                                : TMemoField;
    KunTabkatze                                    : TIntegerField;
    KunTabkatze_anz                                : TIntegerField;
    KunTabhund                                     : TIntegerField;
    KunTabhund_anz                                 : TIntegerField;
    KunTabdvag1                                    : TStringField;
    KunTabdvag2                                    : TStringField;
    KunTabdvag1_bem                                : TStringField;
    KunTabdvag2_bem                                : TStringField;
    KunTabdvag_interesse                           : TIntegerField;
    KunTabdatum_eg                                 : TDateTimeField;
    KunTabdatum_einzug                             : TDateTimeField;
    KunTabdatum_umz                                : TStringField;
    KunTabauswahl                                  : TIntegerField;
    KunTabherkunft                                 : TStringField;
    KunTabletzteaenderung                          : TDateTimeField;
    KunTabMATCHCODE                                : TStringField;
    KunTabKUNDENGRUPPE                             : TIntegerField;
    KunTabSPRACH_ID                                : TIntegerField;
    KunTabGESCHLECHT                               : TStringField;
    KunTabKUNNUM1                                  : TStringField;
    KunTabKUNNUM2                                  : TStringField;
    KunTabLAND                                     : TStringField;
    KunTabNAME3                                    : TStringField;
    KunTabABTEILUNG                                : TStringField;
    KunTabANREDE                                   : TStringField;
    KunTabPOSTFACH                                 : TStringField;
    KunTabPF_PLZ                                   : TStringField;
    KunTabDEFAULT_LIEFANSCHRIFT_ID                 : TIntegerField;
    KunTabGRUPPE                                   : TStringField;
    KunTabTELE1                                    : TStringField;
    KunTabTELE2                                    : TStringField;
    KunTabFUNK                                     : TStringField;
    KunTabINTERNET                                 : TStringField;
    KunTabDIVERSES                                 : TStringField;
    KunTabBRIEFANREDE                              : TStringField;
    KunTabBLZ                                      : TStringField;
    KunTabKTO                                      : TStringField;
    KunTabBANK                                     : TStringField;
    KunTabIBAN                                     : TStringField;
    KunTabSWIFT                                    : TStringField;
    KunTabKTO_INHABER                              : TStringField;
    KunTabDEB_NUM                                  : TIntegerField;
    KunTabKRD_NUM                                  : TIntegerField;
    KunTabNET_SKONTO                               : TFloatField;
    KunTabNET_TAGE                                 : TIntegerField;
    KunTabBRT_TAGE                                 : TIntegerField;
    KunTabWAEHRUNG                                 : TStringField;
    KunTabUST_NUM                                  : TStringField;
    KunTabVERTRETER_ID                             : TIntegerField;
    KunTabPROVIS_PROZ                              : TFloatField;
    KunTabINFO                                     : TMemoField;
    KunTabGRABATT                                  : TFloatField;
    KunTabKUN_KRDLIMIT                             : TFloatField;
    KunTabKUN_LIEFART                              : TIntegerField;
    KunTabKUN_ZAHLART                              : TIntegerField;
    KunTabKUN_PRLISTE                              : TBooleanField;
    KunTabKUN_LIEFSPERRE                           : TBooleanField;
    KunTabLIEF_LIEFART                             : TIntegerField;
    KunTabLIEF_ZAHLART                             : TIntegerField;
    KunTabLIEF_PRLISTE                             : TBooleanField;
    KunTabLIEF_TKOSTEN                             : TFloatField;
    KunTabLIEF_MBWERT                              : TFloatField;
    KunTabPR_EBENE                                 : TIntegerField;
    KunTabBRUTTO_FLAG                              : TBooleanField;
    KunTabMWST_FREI_FLAG                           : TBooleanField;
    KunTabKUN_SEIT                                 : TDateField;
    KunTabKUN_GEBDATUM                             : TDateField;
    KunTabENTFERNUNG                               : TIntegerField;
    KunTabERSTELLT                                 : TDateField;
    KunTabERST_NAME                                : TStringField;
    KunTabGEAEND                                   : TDateField;
    KunTabGEAEND_NAME                              : TStringField;
    KunTabSHOP_KUNDE                               : TBooleanField;
    KunTabSHOP_ID                                  : TIntegerField;
    KunTabSHOP_KUNDE_ID                            : TIntegerField;
    KunTabSHOP_CHANGE_FLAG                         : TIntegerField;
    KunTabSHOP_DEL_FLAG                            : TBooleanField;
    KunTabSHOP_PASSWORD                            : TStringField;
    KunTabUSERFELD_01                              : TStringField;
    KunTabUSERFELD_02                              : TStringField;
    KunTabUSERFELD_03                              : TStringField;
    KunTabUSERFELD_04                              : TStringField;
    KunTabUSERFELD_05                              : TStringField;
    KunTabUSERFELD_06                              : TStringField;
    KunTabUSERFELD_07                              : TStringField;
    KunTabUSERFELD_08                              : TStringField;
    KunTabUSERFELD_09                              : TStringField;
    KunTabUSERFELD_10                              : TStringField;
    KunTabANREDE1                                  : TStringField;
    KunTabANREDE2                                  : TStringField;
    KunTabGEB_DAT_1                                : TDateField;
    KunTabGEB_DAT_2                                : TDateField;
    FirmaTabSMSAbsenderName                        : TStringField;
    MitarbeiterAktivTabidmitarbeiter               : TIntegerField;
    MitarbeiterAktivTabaktiv                       : TIntegerField;
    MitarbeiterAktivTablogin                       : TStringField;
    MitarbeiterAktivTabMA_ID                       : TIntegerField;
    MitarbeiterAktivTabMA_NUMMER                   : TStringField;
    MitarbeiterAktivTabLOGIN_NAME                  : TStringField;
    MitarbeiterAktivTabANZEIGE_NAME                : TStringField;
    MitarbeiterAktivTabUSER_PASSWORD               : TStringField;
    MitarbeiterTabzweig                            : TStringField;
    ZweigTab                                       : TZMySqlQuery;
    ZweisDS                                        : TDataSource;
    ZweigTabrecID                                  : TIntegerField;
    ZweigTabNr                                     : TStringField;
    ZweigTabKurzName                               : TStringField;
    ZweigTabName                                   : TStringField;
    ZweigTabStrasse                                : TStringField;
    ZweigTabPLZ                                    : TStringField;
    ZweigTabOrt                                    : TStringField;
    MitarbeiterAktivTabKontaktMobil                : TStringField;
    MitarbeiterAktivTabKontaktEmail                : TStringField;
    MitarbeiterTabKontaktEmail                     : TStringField;
    ZweigTabTelefon                                : TStringField;
    ZweigTabFax                                    : TStringField;
    ZweigTabmobil                                  : TStringField;
    ZweigTabName2                                  : TStringField;
    ZweigTabName3                                  : TStringField;
    MitarbeiterAktivTabKontaktTelefon              : TStringField;
    MitarbeiterTabKontaktTelefon                   : TStringField;
    ZweigTabemail                                  : TStringField;
    MitarbeiterTabObjektKontingent                 : TIntegerField;
    MitarbeiterAktivTabObjektKontingent            : TIntegerField;
    MitarbeiterTabname1                            : TStringField;
    MitarbeiterTabvorname1                         : TStringField;
    MitarbeiterAktivTabname1                       : TStringField;
    MitarbeiterAktivTabvorname1                    : TStringField;
    MitarbeiterTabKontaktMobil                     : TStringField;
    MitarbeiterAktivTabKalenderfarbe               : TFloatField;
    MitarbeiterAktivTabKurz                        : TStringField;
    MitarbeiterAktivTabzweig                       : TStringField;
    MitarbeiterTabANREDE1                          : TStringField;
    ArtInfoTabINFO                                 : TMemoField;
    ArtInfoTabBESTNUM                              : TStringField;
    ArtInfoTabMENGE_LIEF                           : TFloatField;
    ArtInfoTabMENGE_SOLL                           : TFloatField;
    ArtInfoTabJID                                  : TIntegerField;
    ArtInfoTabALTTEIL_FLAG                         : TStringField;
    ArtInfoTabNO_RABATT_FLAG                       : TStringField;
    ArtInfoTabNO_PROVISION_FLAG                    : TStringField;
    ArtInfoTabNO_BEZEDIT_FLAG                      : TStringField;
    ArtInfoTabNO_VK_FLAG                           : TStringField;
    ArtInfoTabNO_EK_FLAG                           : TStringField;
    ArtInfoTabSN_FLAG                              : TStringField;
    MitarbeiterTabtel_gesch                        : TStringField;
    MitarbeiterTabtel_privat                       : TStringField;
    MitarbeiterTabtel_mobil                        : TStringField;
    MitarbeiterTabANZEIGE_NAME                     : TStringField;
    MitarbeiterTabemail                            : TStringField;
    JourTabObjnr                                   : TIntegerField;
    tabLizenz                                      : TZMySqlQuery;
    tabLizenzStufe                                 : TIntegerField;
    tabLizenzgebuehr                               : TFloatField;
    tabLizenzprovisionsanteil                      : TFloatField;
    tabLizenztext                                  : TStringField;
    tabLizenzrecid                                 : TIntegerField;
    LizenzstufeTab                                 : TZMySqlQuery;
    LizenzstufeTabrecid                            : TIntegerField;
    LizenzstufeTabmitarbeiter                      : TIntegerField;
    LizenzstufeTabstufe                            : TIntegerField;
    LizenzstufeTabeintritt                         : TDateField;
    LizenzstufeTabaustritt                         : TDateField;
    MitarbeiterTabplz                              : TStringField;
    KunTabplz                                      : TStringField;
    db2                                            : TZMySqlDatabase;
    Transact2                                      : TZMySqlTransact;
    tabData                                        : TZMySqlQuery;
    tabDatarec_id                                  : TIntegerField;
    tabDatadata                                    : TBlobField;
    tabuni2                                        : TZMySqlQuery;
    SQLQuery1                                      : TSQLQuery;
    SQLConnection1                                 : TSQLConnection;
    FirmaTabAnfragenSicherung: TStringField;

    PROCEDURE DM1Create(Sender                     : TObject);
    PROCEDURE JourTabCalcFields(DataSet            : TDataSet);
    PROCEDURE KGRTabBeforePost(DataSet             : TDataSet);
    PROCEDURE FirBankTabBeforePost(DataSet         : TDataSet);
    PROCEDURE DataModuleDestroy(Sender             : TObject);
    PROCEDURE NummerTabNewRecord(DataSet           : TDataSet);
    PROCEDURE ShopOrderStatusTabBeforePost(DataSet : TDataSet);
    PROCEDURE ZBatchSql1BeforeExecute(Sender       : TObject);
    PROCEDURE Transact1AfterBatchExec(Sender       : TObject; VAR Res : Integer);
    PROCEDURE ZBatchSql1AfterExecute(Sender        : TObject);
    PROCEDURE FirmaTabCalcFields(DataSet           : TDataSet);
    FUNCTION CaoSecurityFindUser(UserName          : STRING; VAR MA_ID, GRUPPE_ID : Integer) : Boolean;
    PROCEDURE CaoSecurityLoadGruppeRechte(Sender   : TObject);
    PROCEDURE CaoSecurityLoadUserRechte(Sender     : TObject);
    //function CaoSecurityIsFreeLock(ModulID       : String; SatzID: Integer): Boolean;
    FUNCTION CaoSecurityReleaseLock(ModulID        : STRING; SatzID : Integer) : Boolean;
    FUNCTION CaoSecuritySetLock(ModulID            : STRING; SatzID : Integer) : Boolean;
    PROCEDURE WgrTabNewRecord(DataSet              : TDataSet);
  PRIVATE
    { Private-Deklarationen }
    InNewNummer                                    : Boolean;
  PUBLIC
    { Public-Deklarationen }
    CAO_SN                                         : STRING;

    //user                                         : String;
    view_user, // Username, der in Formularen und zur Anzeige verwendet
    // wird, dieser wird aus der Tabelle MITARBEITER ermittelt
//ntuser,
    comp, email                                    : STRING[100];
    Zweig                                          : STRING[3];
    MitarbeiterID                                  : Integer; // Eindeutige Mitarbeiternummer (ID)
    MitarbeiterLizenzstufe                         : Integer ;
    MitarbeiterName                                : STRING[30];
    MitarbeiterKurz                                : STRING[5];
    MitarbeiterAnrede                              : STRING[30];
    MitarbeiterGruppeID                            : Integer;
    MitarbeiterGruppe                              : STRING[30];
    MitarbeiterObjGruppe                           : STRING[30];
    MitarbeiterObjGruppeID                         : Integer;
    LeitWaehrung                                   : STRING;
    LandK2                                         : STRING[2];
    TerminZeitaufloesung                           : Integer;
    TerminZeitStart                                : ttime;
    TerminZeitEnde                                 : ttime;

    MainDir                                        : STRING;
    BackupDir                                      : STRING;
    DTADir                                         : STRING;
    TmpDir                                         : STRING;
    LogDir                                         : STRING;
    ExportDir                                      : STRING;
    ImportDir                                      : STRING;
    DokumentenDir                                  : STRING;

    C2Color                                        : TColor; // Farbe f. 2. Textzeile in Gittern
    EditColor                                      : TColor;

    // zum Cachen der Funktion Calcleitwaehrung
    // sonst muß bei jeder Berechnung in der Registry
    // der Umrechnungskurs geholt werden
    CacheLastWaehrung                              : STRING;
    CacheLastKurs                                  : Double;
    LastVertrID                                    : Integer;
    LastVertrProz                                  : Double;
    SQLLog                                         : Boolean;
    RestoreRun                                     : Boolean;
    TermID                                         : Integer; // Eindeutige ID der Arbeitsstation
    TermIDStr                                      : STRING;
    MwStTab                                        : ARRAY[0..3] OF Double; // Globale MwSt-Tabelle
    DefMwSt                                        : Double; // Default MehrWertSteuer in %
    DefMwStCD                                      : Integer; // Verweis auf einen der Einträge in MwStTab
    AnzPreis                                       : Integer; // max. Preis (VK1-VK5)
    GCalcFaktorTab                                 : ARRAY[1..6] OF Double; // Globale Kalkulationsfaktoren
    MandantTab                                     : ARRAY OF MandantRec;
    AktMandant                                     : STRING;
    MandantOK                                      : Boolean;
    USE_KFZ                                        : Boolean;
    AtrisEnable                                    : Boolean; // True, wenn Atris-Pfad gesetzt ist
    AtrisPfad                                      : STRING;
    DefSpracheID                                   : Integer;
    DefSprachCode                                  : STRING;
    BR_RUND_WERT                                   : Integer; // ganze Cent !!!   z.B. 5 = immer auf ganz 5 Cent runden
    BR_SUM_RUND_WERT                               : Integer; // dito, jedoch für die Belegerstellung
    EK_NACHKOMMA                                   : Integer;
    VK_NACHKOMMA                                   : Integer;
    BLZ_LEN                                        : Integer;
    EK_DFormat                                     : STRING;
    EK_EFormat                                     : STRING;
    VK_DFormat                                     : STRING;
    VK_EFormat                                     : STRING;
    WgrFaktorCache                                 : tWgrFaktorCache;
    MahnFrist                                      : ARRAY[1..5] OF Integer; // Anz. Tage als frist pro Mahnstufe
    PLZ_VERSION,
    BLZ_VERSION                                    : Integer;
    DisplayDLL                                     : STRING;
    IsLinux                                        : Boolean;
    LogLevel                                       : Integer;
    UseNTUserName                                  : Boolean;
    DefaultUserName                                : STRING;
    DefaultPassword                                : STRING;
    Lang_2                                         : STRING; // Ländercode für Sprache z.B. de, fr, es, en

    PROCEDURE AddLog(Typ, s : STRING);
    PROCEDURE ReadMandanten(App : STRING);
    PROCEDURE SaveMandanten;
    FUNCTION GetMandant(Name : STRING; VAR Daten : MandantRec) : Boolean;
    FUNCTION OpenMandant(NewMandant, App : STRING; save : boolean) : boolean;
    PROCEDURE NewMandant(Daten : MandantRec);
    PROCEDURE DeleteMandant(Name : STRING);
    PROCEDURE InitMandantAfterOpen;
    FUNCTION UpdateDatabase(Data : tJvStrHolder; VAR Warnings, Errors : Integer; LogFileName : STRING) : boolean;
    FUNCTION GetDBUserRechte(AktUser : Boolean; User, Secret : STRING) : tSDBUserRechte;

    FUNCTION BucheKasse(Datum      : tDateTime;
                        Quelle     : Integer;
                        Journal_ID : Integer;
                        BelNum     : STRING;
                        GKonto     : Integer;
                        Skonto     : Double;
                        Betrag     : Double;
                        Text       : STRING) : Boolean;

    // Liefert True zurück, wenn die Bankverbindung ok ist
    FUNCTION CheckBankverbindung(addr_id : integer) : boolean;

    // Liefert die Bankverbindung zurück, wenn ok
    FUNCTION GetBankverbindung(addr_id     : integer;
                               VAR BLZ     : Integer;
                               VAR KTO     : STRING;
                               VAR Inhaber : STRING) : boolean; // True wenn ok

    FUNCTION GetLieferant(addr_id : integer; VAR Info : STRING) : boolean; // True wenn ok

    FUNCTION IncNummer(Quelle : Integer) : Int64;
    FUNCTION IncNummerStr(Quelle : Integer) : STRING;
    FUNCTION GetNummerFormat(Quelle : Integer) : STRING;

    // Funktionen für SQL-Registry
    PROCEDURE WriteString(Key, Name, Value : STRING);
    FUNCTION ReadString(Key, Name, Default : STRING) : STRING;
    PROCEDURE WriteBoolean(Key, Name : STRING; Value : Boolean);
    FUNCTION ReadBoolean(Key, Name : STRING; Default : Boolean) : Boolean;
    PROCEDURE WriteInteger(Key, Name : STRING; Value : Integer);
    FUNCTION ReadInteger(Key, Name : STRING; Default : Integer) : Integer;
    FUNCTION ReadInteger2(Key, Name : STRING; Default : Integer) : Integer;
    PROCEDURE WriteDouble(Key, Name : STRING; Value : Double);
    FUNCTION ReadDouble(Key, Name : STRING; Default : Double) : Double;
    FUNCTION ReadLongString(Key, Name, Default : STRING) : STRING;
    PROCEDURE WriteLongString(Key, Name, Value : STRING);

    // USER-SETTINGS
    PROCEDURE WriteStringU(Key, Name, Value : STRING);
    FUNCTION ReadStringU(Key, Name, Default : STRING) : STRING;
    PROCEDURE WriteBooleanU(Key, Name : STRING; Value : Boolean);
    FUNCTION ReadBooleanU(Key, Name : STRING; Default : Boolean) : Boolean;
    PROCEDURE WriteIntegerU(Key, Name : STRING; Value : Integer);
    FUNCTION ReadIntegerU(Key, Name : STRING; Default : Integer) : Integer;
    PROCEDURE WriteDoubleU(Key, Name : STRING; Value : Double);
    FUNCTION ReadDoubleU(Key, Name : STRING; Default : Double) : Double;
    FUNCTION ReadLongStringU(Key, Name, Default : STRING) : STRING;
    PROCEDURE WriteLongStringU(Key, Name, Value : STRING);

    // Funktionen zum lesen und wiederherstellen von Tabellenlayouts (tDBGrid)
    FUNCTION ReadLayout(Key, Name : STRING; VAR Data : tStream; Version : Integer = 0) : Boolean;
    PROCEDURE WriteLayout(Key, Name : STRING; Data : tStream; Version : Integer = 0);

    PROCEDURE GridLoadLayout(VAR Grid : tDBGrid; Sec : STRING; Version : Integer = 0);
    PROCEDURE GridSaveLayout(Grid : tDBGrid; Sec : STRING; Version : Integer = 0);

    // Diverse funktionen zum verbuchen von Rechnungen etc.

    FUNCTION InsertStuecklistenArtikel(JournalID, JournalposID, ArtikelID,
                                       AddrID, BelegArt : Integer;
                                       Menge : Double;
                                       BelegNum : STRING) : Boolean;
    FUNCTION UpdateStuecklistenArtikel(JournalID,
                                       JournalposID,
                                       ArtikelID : Integer;
                                       Menge : Double) : Boolean;

    FUNCTION Buche_Rechnung(Journal_ID : Integer) : STRING; // liefert Rechnungnummer zurück
    FUNCTION Buche_Einkauf(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
    FUNCTION Buche_Gutschrift(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
    FUNCTION Buche_Abrechnung(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
    FUNCTION Buche_Rueckgabe(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
    FUNCTION Buche_Lieferschein(Journal_ID : Integer; Teillief : Boolean; VAR LieferscheinID : Integer) : STRING; // liefert Belegnummer zurück
    FUNCTION Buche_Angebot(Journal_ID : Integer) : STRING; // liefert AGB-Nummer zurück
    FUNCTION Buche_Auftrag(Journal_ID : Integer) : STRING; // liefert Auftrags-Nummer zurück
    FUNCTION Buche_BCKasse(Journal_ID : Integer) : STRING; // liefert BON-Nummer zurück
    FUNCTION Buche_EKBest(Journal_ID : Integer) : STRING; // liefert EK-BST-Nummer zurück

    // Diverse Funktionen zum Stornieren von Vorgängen
    FUNCTION Storno_Einkauf(Journal_ID : Integer) : Boolean; // True, Wenn OK
    FUNCTION Storno_Verkauf(Journal_ID : Integer) : Boolean; // True, Wenn OK
    FUNCTION Storno_Angebot(Journal_ID : Integer) : Boolean; // True, Wenn OK
    FUNCTION Storno_Lieferschein(Journal_ID : Integer) : Boolean; // True, Wenn OK
    FUNCTION Storno_EKBestellung(Journal_ID : Integer) : Boolean; // True, Wenn OK

    // Rechnet beliebige Währung in Leitwährung um
    FUNCTION CalcLeitWaehrung(Betrag : Double; Waehrung : STRING) : Double;

    // Kopiert Belege
    FUNCTION CopyRechnung(Journal_ID, Dest : Integer) : Integer; // Liefert Rec-ID zurück

    FUNCTION GetWGRDefaultKonten(WGR : Integer; VAR EKTO, AKTO : Integer) : Boolean;
    FUNCTION CalcRabGrpPreis(RGID : STRING; PR_Ebene : Integer; VAR Preis : Double) : Boolean;

    FUNCTION UpdateArtikelEdiMenge(JournalTyp, ArtikelID : Integer; MengeDiff : Double) : Boolean;
    FUNCTION UpdateEKBestMenge : Boolean;
    PROCEDURE UpdateArtikelPreis(RechTyp, Artikel_ID, Addr_ID : Integer; Preis : Double);

    // Export-CSV Funktionen
    PROCEDURE ExportCSVDatasetToStream(Stream : TStream;
                                       Dataset : TDataset;
                                       Delimiter : STRING;
                                       Spaltennamen : Boolean = True;
                                       TextInHochKomma : Boolean = True;
                                       DosZeichenSatz : Boolean = False);

    PROCEDURE ExportCSVDatasetToFile(FileName : STRING;
                                     Dataset : TDataset;
                                     Delimiter : STRING;
                                     Append : Boolean;
                                     Spaltennamen : Boolean = True;
                                     TextInHochKomma : Boolean = True;
                                     DosZeichenSatz : Boolean = False);

    PROCEDURE ExportCSVDatasetToExcel(FileName : STRING; Dataset : TDataset; Spaltennamen : Boolean = True);

    // Export-Funktionen
    PROCEDURE ExportDatasetToStream(Stream : TStream; Dataset : TDataset;
                                     Delimiter : STRING; Spaltennamen : Boolean = True; TextInHochKomma : Boolean
                                     = True; DosZeichenSatz : Boolean = False);
    PROCEDURE ExportDatasetToFile(FileName : STRING; Dataset : TDataset;
                                  Delimiter : STRING; Append : Boolean; Spaltennamen : Boolean = True;
                                  TextInHochKomma : Boolean = True; DosZeichenSatz : Boolean = False);
    PROCEDURE ExportDatasetToExcel(FileName : STRING; Dataset : TDataset; Spaltennamen : Boolean = True);
    FUNCTION GetSearchSQL(Felder : ARRAY OF STRING; Suchbegriff : STRING) : STRING;
    FUNCTION GetSearchSQL2(Felder : ARRAY OF STRING; Suchbegriff : STRING) : STRING;
    FUNCTION GetVertreterProv(ID : Integer) : Double;
    PROCEDURE CheckTerminalID;
    //--------------------
    FUNCTION GetArtikelPreis(ArtikelID, KunID, PE : Integer; Brutto : Boolean; Menge : Double; VAR Preis : Double) : Boolean;
    FUNCTION GetWGRCalcFaktor(Wgr, PreisID : Integer; VAR Faktor : Double) : Boolean;
    PROCEDURE StarteNewProgramm(Name, Cmd, Dir : STRING);
    PROCEDURE LockError(Error : Integer);
    FUNCTION ObjektTransfer(Nr : Integer) : Boolean;
    FUNCTION Bildname(s : STRING; Typ : STRING) : STRING;
    // Für Profilerzeugung
    FUNCTION SummeAdd(Quelle, Was : STRING) : STRING;
  END;
  
FUNCTION Berechtigung(was : integer) : Boolean;
FUNCTION GetProjectVersion : STRING;
FUNCTION AddBackSlash(Name : STRING) : STRING;
FUNCTION DateToSql(D : TDateTime) : STRING;
//------------------------------------------------------------------------------
FUNCTION EmailSend(Quelle : STRING; Werte : ARRAY OF ParseRec) : Boolean;
(*
rocedure SendEMaul;
VAR
  R : Array of ParseRec;
begin
  SetLength(R,15);

  R[0].Name := 'EmpfaengerEmail';
  R[0].Wert := 'gerhard.p@geram.de';
  R[1].Name := 'Absendername';
  R[1].Wert := 'wir';
  R[2].Name := 'Absenderadresse';
  R[2].Wert := 'info@geram.de';
  R[3].Name := 'Ueberschrift';
  R[3].Wert := 'test von uns(delphi)';
  R[4].Name := 'Antwortname';
  R[4].Wert := 'Gerhard der große';
  R[5].Name := 'Antwortadresse';
  R[5].Wert := 'gerhard@geram.de';



  R[10].Name := 'Name';
  R[10].Wert := 'Graef';
  R[11].Name := 'Vorname';
  R[11].Wert := 'Stefan';
  R[12].Name := 'Anrede';
  R[12].Wert := 'Herr';
  R[13].Name := 'Strasse';
  R[13].Wert := 'Akeleistr. 1';
  R[14].Name := 'Ort';
  R[14].Wert := 'Gröbenzell';

  EmailSend('Immo_Angebot',R);
end;
*)
FUNCTION SendSMS(Quelle : STRING; Werte : ARRAY OF ParseRec) : Boolean;
FUNCTION DateiSpeichern(S:String;Name:String;Overwrite:boolean=false):Integer;
Procedure DateiLoeschen(S : String);
PROCEDURE DateiZugriffSuchen;


VAR
  DM1                   : TDM1;
  CAO32_DB_CFG          : STRING[12] = 'CAO32DB.CFG'; //GERA
  DokumentenFTPHost,
  DokumentenFTPName,
  DokumentenFTPPasswort : String;

//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
IMPLEMENTATION

USES
  inifiles
  , FileCtrl
{$IFDEF GNUTEXT}
  , gnugettext //OK
{$ENDIF}
  , cao_sqlrechtedlg
  //, cao_register
{$IFDEF COMPILER_D6_UP}
  , Variants
{$ENDIF}
  , CAO_Link
  , shellapi
  , XLSFile
  , CAO_Logging
  , ZExtra
  , cao_progress
  , cao_searchclass
  , ZSqlTypes
  , CAO_Tool1
  , cao_shoptrans
  , IMMO_Terminierung, CAO_MAIN;

{$R *.DFM}

{$IFDEF SMTPLOG}
VAR
  SMTPFile : File;
{$ENDIF SMTPLOG}

CONST
  abc : STRING = '@12345@';

  // aktuell verwendete Datenbank-Version (Struktur/Intern) = 1.10
  // 1.11 1.10.08 + feld Status -
  // wird mit der Version in der SQL-Registry verglichen
  // und das öffnen einer DB mit zu großer DB-Version (Programm älter als DB)
  // abgebrochen
CONST
  DBVersion_Soll = 113;
//------------------------------------------------------------------------------
PROCEDURE DateiZugriffSuchen;
BEGIN
  IF DirectoryExists(DM1.DokumentenDir) THEN
    Dateizugriff := ZugriffDirekt
  ELSE
    BEGIN
      Dateizugriff := ZugriffFTP;
      DokumentenFTPHost     := DM1.ReadString('DOKUMENTEN','FTPHOSTDOKUMENTEN', '');
      DokumentenFTPName     := DM1.ReadString('DOKUMENTEN','FTPBENUTZERDOKUMENTEN', '');
      DokumentenFTPPasswort := DM1.ReadString('DOKUMENTEN','FTPPASSWORTDOKUMENTEN', '');
    END;
end;

//------------------------------------------------------------------------------
FUNCTION DateToSql(D : TDateTime) : STRING;
BEGIN
  Result := FormatDateTime('yyyy-mm-dd hh:nn:ss', D);
END;

{$IFDEF SMTPLOG}
//------------------------------------------------------------------------------
PROCEDURE SchreibeSMTPLOG(TheMessage : TIdMessage);
VAR
  Written : Integer;
BEGIN
  AssignFile(SMTPFile,ExtractFilePath(Application.ExeName)+'SMTPMessage.log');
  try
    Reset(SMTPFile);
  except
    Rewrite(SMTPFile);
  end;
  Seek(SMTPFile,FileSize(SMTPFile));
  BlockWrite(SMTPFile,TheMessage,SizeOf(TheMessage),Written);
  IF (Written <> SizeOf(TheMessage)) THEN
    BEGIN
      MainForm.SBar.Panels[5].Text := 'SMTP Fehler '+TheMessage.Subject;
      MessageBeep(MB_ICONHAND);
      LogForm.AddLog(MainForm.SBar.Panels[5].Text);
    END;
  CloseFile(SMTPFile);
end;
{$ENDIF}


//------------------------------------------------------------------------------
FUNCTION EmailSend(Quelle : STRING; Werte : ARRAY OF ParseRec) : Boolean;

VAR
  F         : TextFile;
  Inhalt, S : STRING;
  I         : Integer;
  html      : TStrings;
  htmlpart  : TIdText;
  txtpart   : TIdText;

  FUNCTION Wert(Such : STRING) : STRING;
  VAR ii : Integer;
  BEGIN
    Result := '';
    FOR II := 0 TO HIGH(Werte) DO
      IF Werte[II].Name = Such THEN
        BEGIN
          Result := Werte[II].Wert;
          exit;
        END;
  END;
BEGIN
  Result := false;
  IF Wert('Kundennr') = '' THEN exit;
  IF NOT ISValidEmail(Wert('EmpfaengerEmail')) THEN exit;

  Quelle := ExtractFileName(Quelle);
  AssignFile(f, ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.txt');
  TRY
    Reset(F);
  EXCEPT
    ShowMessage('Datei ' + ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.txt nicht gefunden!');
    Exit;
  END;
  Inhalt := '';

  WHILE NOT EOF(F) DO
    BEGIN
      try
        ReadLn(F,S);
      except
        ShowMessage('Datei ' + ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.txt lesefehler.');
        Exit;
      end;
      Inhalt := Inhalt + S + #$0D#$0A;
    END;
  CloseFile(f);

  FOR I := 0 TO HIGH(Werte) DO
    BEGIN
      Inhalt := StringReplace(Inhalt, '{' + Werte[I].Name + '}', Werte[I].Wert, [rfReplaceAll, rfIgnoreCase]);
    END;

  DM1.emMessage.Clear;
  DM1.emMessage.ContentType := 'multipart/mixed';
  DM1.emMessage.ReplyTo.Add;
  DM1.emMessage.ReplyTo.Items[0].Name := wert('Antwortname');
  DM1.emMessage.ReplyTo.Items[0].Address := wert('Antwortadresse');
  DM1.emMessage.From.Name := wert('Absendername');
  DM1.emMessage.From.Address := wert('Absenderadresse');
  DM1.emMessage.Subject := Wert('Ueberschrift');
// Form1.emMessage.Sender.Name  := Wert('Sendername');
// Form1.emMessage.Sender.Text := wert('');
// Form1.emMessage.Sender.DisplayName := wert('');
  DM1.emMessage.Sender.Address := Wert('AbsenderAdresse');
  DM1.emMessage.Recipients.EMailAddresses := Wert('EmpfaengerEmail');
  txtPart := TIdText.Create(DM1.emMessage.MessageParts);
  txtPart.Body.Add(Inhalt);


  AssignFile(f, ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.html');
  TRY
    Reset(F);
  EXCEPT
    ShowMessage('Datei ' + ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.html nicht gefunden!');
    Exit;
  END;
  Inhalt := '';

  WHILE NOT EOF(F) DO
    BEGIN
      try
        ReadLn(F, S);
      except
        ShowMessage('Datei ' + ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.html lesefehler.');
        Exit;
      end;
      Inhalt := Inhalt + S + #$0D#$0A;
    END;
  CloseFile(f);

  FOR I := 0 TO HIGH(Werte) DO
    BEGIN
      Inhalt := StringReplace(Inhalt, '{' + Werte[I].Name + '}', Werte[I].Wert, [rfReplaceAll, rfIgnoreCase]);
    END;
  htmlpart := TIdText.Create(DM1.emMessage.MessageParts, html);
  htmlpart.ContentType := 'text/html';
  htmlpart.Body.Add(Inhalt);

  DM1.emSMTP.Host := DM1.FirmaTabeMailSMTP.AsString;
  DM1.emSMTP.Username := DM1.FirmaTabeMailBenutzer.AsString;
  DM1.emSMTP.Password := DM1.FirmaTabeMailPasswort.AsString;

  IF DM1.emSMTP.Host = '' THEN
    BEGIN
      ShowMessage('Der eMail SMTP Hostname fehlt.');
//      Screen.Cursor := crDefault;
      Exit;
    END;

  IF DM1.emSMTP.Username = '' THEN
    BEGIN
      ShowMessage('Der eMail Benutzername fehlt.');
//      Screen.Cursor := crDefault;
      Exit;
    END;

  IF DM1.emSMTP.Password = '' THEN
    BEGIN
      ShowMessage('Das eMail Passwort fehlt.');
//      Screen.Cursor := crDefault;
      Exit;
    END;

  TRY
    DM1.emSMTP.Connect;
  EXCEPT
    ShowMessage('Fehler beim eMailversand! SMTP Host:' + DM1.emSMTP.Host + ' Sendeadresse:' + DM1.emSMTP.Username + ^M + Format(' Fehler %d', [GetLastError]));
//    mInfo.Lines.Add('Fehler beim eMailversand! SMTP Host:' + emSMTP.Host +' Sendeadresse:' + emSMTP.Username + ^M + Format(' Fehler %d', [GetLastError]));
    exit;
  END;

  IF DM1.emSMTP.Connected THEN
    TRY
      {$IFDEF SMTPLOG}
//      SchreibeSMTPLOG(DM1.emMessage);
      {$ELSE}
      DM1.emSMTP.Send(DM1.emMessage);
      {$ENDIF}
    EXCEPT
      DM1.emSMTP.Disconnect;
      LogForm.AddLog('"'+DM1.emMessage.Subject+'";"'+DM1.emMessage.ReplyTo.Items[0].Name+'";"'+DM1.emMessage.ReplyTo.Items[0].Address+'";"'+DM1.emMessage.Recipients.EMailAddresses+'"');
      exit;
    END;

  DM1.emSMTP.Disconnect;
//  mInfo.Lines.Add('    eMail an ' + wert('EmpfaengerEmail') + ' verschickt!');
  Result := true;
  DM1.Uniquery.Sql.Text := 'insert into tbl_kommunikation set Datum=now(),Art=1,KundenID ="' + Wert('Kundennr') + '", ObjektID="' + Wert('Objektnr') + '",Ursprung ="' + Wert('Ursprung') + '",Text="' + txtpart.Body.Text + '"';
  DM1.Uniquery.ExecSql;
END;
//------------------------------------------------------------------------------

FUNCTION SendSMS(Quelle : STRING; Werte : ARRAY OF ParseRec) : Boolean;

VAR
  F         : TextFile;
  Inhalt, S : STRING;
  I         : Integer;
  html      : TStrings;
  htmlpart  : TIdText;
  txtpart   : TIdText;
  adr, url  : STRING;

  FUNCTION ValidPhoneNr(PhoneNr : STRING) : STRING;
  VAR II : Integer;
  BEGIN
    Result := '';
    FOR II := 1 TO Length(PhoneNr) DO
      IF PhoneNr[II] IN ['+', '0'..'9'] THEN
        Result := Result + PhoneNr[II];
  END;


  FUNCTION Wert(Such : STRING) : STRING;
  VAR ii : Integer;
  BEGIN
    Result := '';
    FOR II := 0 TO HIGH(Werte) DO
      IF Werte[II].Name = Such THEN
        BEGIN
          Result := Werte[II].Wert;
          exit;
        END;
  END;

  FUNCTION ValidURL(Such : STRING) : STRING;
  VAR ii : Integer;
  BEGIN
    Result := '';
    FOR II := 1 TO Length(Such) DO
      BEGIN
        IF ((Such[II] = #$0D) OR (Such[II] = #$0A)) THEN continue;
        IF (Such[II] = ' ') THEN
          Result := Result + '+'
        ELSE
          Result := Result + Such[II];
      END;
//      StringReplace(Result,' ','+',[rfReplaceAll]);
  END;

BEGIN
  Result := false;
  IF Wert('EmpfaengerSMSNr') = '' THEN exit; // keine telefonnummer angegeben
  IF Wert('Kundennr') = '' THEN exit;

  Quelle := ExtractFileName(Quelle);
  AssignFile(f, ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.txt');
  TRY
    Reset(F);
  EXCEPT
    ShowMessage('Datei ' + ExtractFilePath(Application.ExeName) + 'Vorlagen\' + Quelle + '.txt nicht gefunden!');
    Exit;
  END;
  Inhalt := '';

  WHILE NOT EOF(F) DO
    BEGIN
      ReadLn(F, S);
      Inhalt := Inhalt + S + ' ';
    END;
  CloseFile(f);
  FOR I := 1 TO HIGH(Werte) DO
    BEGIN
      Inhalt := StringReplace(Inhalt, '{' + Werte[I].Name + '}', Werte[I].Wert, [rfReplaceAll, rfIgnoreCase]);
    END;

  TRY
    adr := DM1.FirmaTabSMSURL.AsString; //'https://gateway.fitsms.de/sms/http2sms.jsp';

    url := adr +
      '?username=' + DM1.FirmaTabSMSBenutzerName.AsString +
      '&password=' + DM1.FirmaTabSMSPasswort.AsString +
      '&type=text' +
      '&to=' + ValidPhoneNr(Wert('EmpfaengerSMSNr')) +
      '&content=' + Inhalt +
      '&from=' + copy(DM1.FirmaTabNAME1.AsString, 1, 10);
    IF Wert('Termin') > '' THEN
      url := url + '&termin=' + wert('Termin');
 // ((  StringReplace(url,' ','+',[rfReplaceAll]);
    URL := validURL(url);
    DM1.AddLog('SMSok', URL);
    TRY
      Adr := DM1.IdHTTP1.Get(ValidURL(url));
    EXCEPT
      DM1.AddLog('SMS', URL);
    END;
    IF Pos('SUCCESS', Adr) > 0 THEN
      BEGIN
        result := TRUE;
        DM1.Uniquery.Sql.Clear;
        DM1.Uniquery.Sql.Add('insert into tbl_kommunikation set Datum=now(),Art=2');
        IF Wert('Kundennr') <> '' THEN
          DM1.Uniquery.Sql.Add(',KundenID ="' + Wert('Kundennr') + '"');
        IF Wert('Objektnr') <> '' THEN
          DM1.Uniquery.Sql.Add(',ObjektID="' + Wert('Objektnr') + '"');
        IF Wert('Ursprung') <> '' THEN
          DM1.Uniquery.Sql.Add(',Ursprung ="' + Wert('Ursprung') + '"');
        DM1.Uniquery.Sql.Add(',text="' + url + '"');
        DM1.Uniquery.ExecSql;
      END
    ELSE
      BEGIN
        DM1.AddLog('SMS', URL);
        DM1.AddLog('SMS', Adr);
        Result := False;
      END;
  EXCEPT
    result := false;
  END;

END;
(*
///------------------------------------------------------------------------------
Function SendSMS(ZielNr,SMSText,KundenID,ObjektNr,Ursprung : String): Boolean;
var
  adr,url : String;

  FUNCTION ValidPhoneNr(PhoneNr : String) : String;
  VAR I : Integer;
  BEGIN
    Result := '';
    FOR I := 1 TO Length(PhoneNr) DO
      IF PhoneNr[I] IN ['+','0'..'9'] THEN
        Result := Result + PhoneNr[I];
  END;

BEGIN
  try
    adr := DM1.FirmaTabSMSURL.AsString; //'https://gateway.fitsms.de/sms/http2sms.jsp';

    url := adr+
           '?username='+DM1.FirmaTabSMSBenutzerName.AsString+
           '&password='+DM1.FirmaTabSMSPasswort.AsString+
           '&type=text'+
           '&to='+ValidPhoneNr(ZielNr)+
           '&content='+SMSText+
           '&from='+DM1.FirmaTabNAME1.AsString;

    //result := DM1.IdHTTP1.Get(url) <> '';
    //adr := 'http://gateway.fitsms.de/sms/http2sms.jsp';
    try
      Adr := DM1.IdHTTP1.Get(url);
    except
      DM1.AddLog('SMS',Adr+' | '+URL);
    end;
    IF Pos('Correct',Adr) > 0 THEN
      BEGIN
        result := TRUE;
        DM1.Uniquery.Sql.Text := 'insert into tbl_kommunikation set Datum=now(),Art=2,KundenID ="'+KundenID+'", ObjektID="'+ObjektNr+'",Ursprung ="'+Ursprung+'"';
        DM1.Uniquery.ExecSql;
      END
    ELSE
      Result := False;
  except
    result := false;
  end;
END;
*)
///------------------------------------------------------------------------------

FUNCTION AddBackSlash(Name : STRING) : STRING;
BEGIN
  IF (Name <> '') THEN
    IF Name[Length(Name)] <> '\' THEN
      Name := Name + '\';

  Result := Name;
END;

PROCEDURE TDM1.DM1Create(Sender : TObject);
VAR
  P : PChar;
  Size : DWord;
  S : STRING;
  I : Integer;
BEGIN
{$IFDEF GNUTEXT}
  TRY
    TranslateComponent(self);
  EXCEPT
  END;
{$ENDIF}

  CAO_SN := '12345-67890-GERA-09876-54321';

  LogLevel := 10;

  Lang_2 := 'de';

  InNewNummer := False;
  SQLLog := True;
  RestoreRun := False;

  IsLinux := False;

  IF ParamCount > 0 THEN
    BEGIN
      FOR i := 0 TO ParamCount DO
        BEGIN
          S := Uppercase(ParamStr(I));
          IF (length(S) > 0) AND (S[1] = '/') THEN
            delete(S, 1, 1);
          IF (length(S) > 0) AND (S[1] = '-') THEN
            delete(S, 1, 1);

          IF S = 'LINUX' THEN
            IsLinux := True;
        END;
    END;

  DB1.Connected := False;
  DB1.Database := '';
  DB1.Host := '';
  DB1.Login := '';
  DB1.Password := '';
  DB1.LoginPrompt := False;

  TermID := -1;
  TermIDStr := '';
  AtrisEnable := False;

  // Rechnernamen ermitteln
  size := 1024;
  p := StrAlloc(Size);
  windows.getcomputername(P, Size);
  comp := p;
  strdispose(p);

  MainDir := ExtractFilePath(Paramstr(0));
  LogDir := MainDir + 'LOG\';
  ForceDirectories(LogDir);
  AnzPreis := 5; //default = VK5
  USE_KFZ := False;
  DefMwSt := 19; // in %
  DefMwStCD := 2;
  DefSpracheID := 2; // Deutsch
  DefSprachCode := 'de';
  LastVertrID := -1;
  LastVertrProz := 0;
  AktMandant := '';
  MandantOK := False;

  EK_NACHKOMMA := 3;
  VK_NACHKOMMA := 2;

  WgrFaktorCache.Wgr := -1; // Cache ungültig

  SetLength(MandantTab, 0);
  ReadMandanten(application.name);

  //MessageDlg (abc,mtinformation,[mbok],0);

//   DisplayTrace('GetCurrentLanguage');
{$IFDEF GNUTEXT}
  Lang_2 := GetCurrentLanguage;
  Lang_2 := LowerCase(Copy(Lang_2, 1, 2));
{$ELSE}
  Lang_2 := 'DE';
{$ENDIF}
  //   DisplayTrace(Lang_2);

  END;

//------------------------------------------------------------------------------
// Öffnet einen bestehenden Mandanten, wenn dieser nicht existiert wird versucht
// die DB und die Tabellen anzulegen, dabei werden die akt. Benutzerrechte
// geprüft und bei zu wenigen Rechten abgebrochen.
// Weiterhin wird die Version der Tabellenstruktur geprüft und ggf. aktualisiert.
// auch hierbei werden die Benutzerrrechte gecheckt.
// bei einer zu neuen Tabellenversion wird das öffnen des Mandanten abgebrochen.
//------------------------------------------------------------------------------

FUNCTION TDM1.OpenMandant(NewMandant, App : STRING; save : boolean) : boolean;
VAR
  Mandant : MandantRec;
  IniName : STRING;
  MyIni : tIniFile;
  NewDB : Boolean;
  I : Integer;
  DSTab : ARRAY OF Boolean;
  V : Double;
  S : STRING;
  Warn : Integer;
  Error : Integer;
  ConOK,
    DBOK : Boolean;
  Res : tSDBUserRechte;
  //    ST      : tDateTime;
  FehlTab : STRING; // Liste der fehlenden Tabellen

  PROCEDURE SqlRechteFehler(Txt : STRING; Rechte : tSDBUserRechte);
  VAR
    MySqlRechteDlg : tSqlRechteDlg;
  BEGIN
    MySqlRechteDlg := tSqlRechteDlg.Create(Self);
    TRY
      MySqlRechteDlg.ShowDlg(Txt, Rechte);
    FINALLY
      MySqlRechteDlg.Free;
    END;
  END;

  PROCEDURE MsgNoSQLRights(Rechte : tSDBUserRechte);
  BEGIN
    SqlRechteFehler
      (_('Sie verfügen über zu wenige Benutzerrechte auf dem MySQL-Server' +
      #13#10 +
      'um die Tabellenstruktur für den akt. Mandanten zu aktualisieren.' + #13#10
      +
      'Sie benötigen mind. folgende Rechte :' + #13#10 +
      'SELECT, INSERT, UPDATE, DELETE, CREATE, DROP, ALTER und INDEX' + #13#13 +
      'Bitte loggen Sie sich mit genügend Rechten erneut ein.'), Rechte);
  END;

  FUNCTION ReadString(Key, Name, Default : STRING) : STRING;
  VAR
    Tab : STRING;
  BEGIN
    IF NOT DB1.Connected THEN
      exit;

    UniQuery.Close;
    UniQuery.Sql.Text := 'show tables like "registry"';
    UniQuery.Open;
    IF UniQuery.RecordCount = 1 THEN
      Tab := UniQuery.Fields[0].AsString
    ELSE
      Tab := 'REGISTRY';

    UniQuery.Close;
    UniQuery.Sql.Text := 'select * from ' + Tab + ' ' +
      'where MAINKEY=:KEY and NAME=:NAME';
    UniQuery.ParamByName('KEY').AsString := Key;
    UniQuery.ParamByName('NAME').AsString := Name;
    UniQuery.Open;

    IF UniQuery.RecordCount > 0 THEN
      Result := UniQuery.FieldByName('VAL_CHAR').AsString
    ELSE
      Result := Default;
    UniQuery.Close;
  END;

BEGIN
  Result := False;
  NewDB := False;
  IF GetMandant(NewMandant, Mandant) THEN
    BEGIN
      // offene Datasets merken
      SetLength(DSTab, DB1.DatasetCount);
      FOR i := 0 TO DB1.DatasetCount - 1 DO
        DSTab[i] := tDataset(DB1.Datasets[i]).Active;

      DB1.Disconnect;

{$IFDEF AVE}
      IF (Assigned(SSHForm)) AND (SSHForm.EndeBtn.Enabled) THEN
        BEGIN
          SSHForm.Show;
          SSHForm.Logout;
          ST := Now();
          REPEAT
            Application.ProcessMessages;
          UNTIL (SSHForm.EndeBtn.Enabled = False) OR
            (Now > ST + (1 / 24 / 60 / 60) * 30); // max. 30 Sek.
        END;
{$ENDIF}

      GetMandant(NewMandant, Mandant);

      DB1.Host := Mandant.Server;
      DB1.Login := Mandant.User;
      DB1.Password := Mandant.Pass;
      DB1.Database := Mandant.DB;

{$IFDEF AVE}
      IF Mandant.UseSSH THEN
        BEGIN
          IF NOT Assigned(SSHForm) THEN
            SSHForm := tSSHForm.Create(Self);
          SSHForm.Show;
          SSHForm.StartBtnClick(Self);
          ST := Now();
          REPEAT
            Application.ProcessMessages;
          UNTIL ((SSHForm.StartBtn.Enabled = False) AND
            (Now > ST + (1 / 24 / 60 / 60) * 15)) OR
            (Now > ST + (1 / 24 / 60 / 60) * 60); // max. 30 Sek.
        END;
{$ENDIF}

      //NEU
      DB1.Port := IntToStr(Mandant.Port);
      //if Mandant.UseNTUserName then DB1.Login :=User;
      //DB1.LoginPrompt :=Mandant.ShowLoginDlg;

      ConOK := False;
      DBOK := False;
      TRY
        DB1.Connect;

        //if DB1.LoginPrompt then User :=DB1.Login else User :=NTUser;

        ConOK := True;
        DBOK := True;

{$IFDEF AVE}
        IF (Assigned(SSHForm)) THEN
          SSHForm.Hide;
{$ENDIF}

      EXCEPT
        ON E : Exception DO
          BEGIN
            IF (Pos('UNBEKANNTE DATENBANK', Uppercase(E.Message)) = 1) OR
              (Pos('UNKNOWN DATABASE', Uppercase(E.Message)) = 1) THEN
              BEGIN
                IF MessageDlg(_('Die Datenbank scheint noch nicht zu ' +
                  'existieren.' + #13#10 +
                  'Wollen Sie die Datenbank erstellen ?'),
                  mtconfirmation, [mbyes, mbno], 0) = mryes THEN
                  BEGIN
                    TRY
                      DB1.CreateDatabase(Mandant.DB); // DB anlegen

                      DBOK := True;
                      DB1.Connect;
                      NewDB := True;
                      ConOK := True;
                    EXCEPT
                      //MessageDlg ('Fehler beim erzeugen der Datenbank !',
                      //            mterror,[mbok],0);
                    END;
                  END
                ELSE
                  exit;
              END
            ELSE
              BEGIN
                MessageDlg(_('Fehler beim verbinden zum MySQL-Server.' +
                  #13#10#13#10 +
                  'Meldung :') + #13#10 + E.Message, mterror, [mbok], 0);
                exit;
              END;
          END;
      END;

      IF NOT DBOK THEN
        BEGIN
          MessageDlg(_('Die Datenbank für diesen Mandant konnte nicht ' +
            'erstellt werden !' + #13#10 +
            'Bitte prüfen Sie die Einstellungen und ob Sie über ' +
            'ausreichende Rechte' + #13#10 +
            'zum erstellen einer DB auf diesem Server verfügen.'),
            mterror, [mbok], 0);
          exit;
        END;

      IF NOT ConOK THEN
        BEGIN
          MessageDlg(
            _('Der gewünschte Mandant konnte nicht geöffnet werden !' + #13#10 +
            'Überprüfen Sie die Verbindung zum Server, Benutzernamen' + #13#10 +
            'und Paßwort.'),
            mterror, [mbok], 0);

          exit;
        END;

      IF NOT NewDB THEN
        BEGIN
          // prüfen, ob Tabellen vorhanden sind, wenn nicht, dann neu erzeugen
          TRY
            uniquery.close;
            uniquery.sql.text := 'SHOW TABLES';
            uniquery.open;
            NewDB := (UniQuery.RecordCount = 0) OR
              (
              (UniQuery.RecordCount = 1) AND
              (UniQuery.FieldByName('Tables_in_' + DB1.Database).AsString =
              'UCHECK')
              );

            IF NOT NewDB THEN
              BEGIN
                UniQuery.First;
                FehlTab := '';
                FOR i := 0 TO length(CAO_TABELLEN_V109) - 1 DO
                  BEGIN
                    IF NOT UniQuery.Locate('Tables_in_' + DB1.Database,
                      CAO_TABELLEN_V109[i], []) THEN
                      BEGIN
                        IF length(FehlTab) > 0 THEN
                          FehlTab := FehlTab + ', ';
                        FehlTab := FehlTab + CAO_TABELLEN_V109[i];
                      END;
                  END;
              END;
          EXCEPT

          END;
          Uniquery.Close;
        END;

      TRY
        IF NewDB THEN // neue DB wurde erstellt
          BEGIN
            Res := GetDBUserRechte(True, '', '');
            IF (urSelect IN Res) AND (urInsert IN Res) AND
              (urUpdate IN Res) AND (urDelete IN Res) AND
              (urCreate IN Res) AND (urAlter IN Res) AND
              (urDrop IN Res) AND (urIndex IN Res) THEN
              BEGIN
                // User hat die benötigten Rechte ! Jetzt Tabellen anlegen ...

                IF NOT UpdateDatabase(CreateMandantStr, Warn, Error,
                  'db_create_110') {//CreateMandantSql} THEN
                  MessageDlg('Fehler beim erstellen der Tabellen !',
                    mterror, [mbok], 0);
              END
            ELSE
              BEGIN
                SqlRechteFehler(
                  _('Sie verfügen über zu wenige Benutzerrechte auf ' +
                  'dem MySQL-Server um die Tabellenstruktur für einen ' +
                  'neuen Mandanten (DB) zu erzeugen.' + #13#10#13#10 +
                  'Sie benötigen mind. folgende Rechte :' + #13#10#13#10 +
                  'SELECT, INSERT, UPDATE, DELETE, CREATE, DROP, ' +
                  'ALTER und INDEX' + #13#13#13#10 +
                  'Bitte loggen Sie sich mit genügend Rechten erneut ein.'),
                  Res);

                DB1.Disconnect;
                exit;
              END;
          END;
      EXCEPT
        MessageDlg(_('Fehler beim erstellen der Tabellen !'), mterror, [mbok],
          0);
        exit;
      END;

      TRY
        IF NewDB THEN // wenn neue Tabellen angelegt, dann evt. Formulare importieren ...
          IF (fileexists(extractfilepath(paramstr(0)) + 'formulare.cao')) AND
            (MessageDlg(_('Sollen die Standardformulare installiert werden ?'),
            mtconfirmation, mbyesnocancel, 0) = mryes) THEN
            BEGIN
              ZBatchSQL1.Sql.Clear;
              ZBatchSQL1.Sql.LoadFromFile(extractfilepath(paramstr(0)) +
                'formulare.cao');
              ProgressForm.Init(_('Formulare installieren ...'));
              ZBatchSQL1.ExecSql;
              ProgressForm.Stop;
            END;
      EXCEPT
        MessageDlg(_('Fehler beim anlegen der Formulare !'), mterror, [mbok],
          0);
      END;

      ProgressForm.Stop;

      TRY
        S := ReadString('MAIN', 'DB_VERSION', '0.00');

        IF DecimalSeparator <> '.' THEN
          WHILE (Pos('.', S) > 0) DO
            S[Pos('.', S)] := DecimalSeparator;
        V := StrToFloat(S);
      EXCEPT
        //           V :=-1; // zur Sicherheit, damit die nä. Updateschritte nicht ausgeführt werden !!!

        MessageDlg(
          _('Die Datenbankversion konnte nicht ermittelt werden !' + #13#10 +
          'Der ausgewählte Mandant wurde nicht geöffnet.'),
          mterror, [mbok], 0);

        DB1.Disconnect;
        exit;
      END;

      //Check ob DB zu alt (Version < 1.09
      TRY
        S := ReadString('MAIN', 'DB_VERSION', '0.00');
        IF DecimalSeparator <> '.' THEN
          WHILE (Pos('.', S) > 0) DO
            S[Pos('.', S)] := DecimalSeparator;
        V := StrToFloat(S) * 100;

        IF V < 109 THEN
          BEGIN
            MessageDlg(_('Die Datenbank ist zu alt (<1.09).' + #13#10 +
              'Bitte Installieren Sie zunächst die Version 1.2.5.x' + #13#10 +
              'von GERA-Faktura und dann erst eine Version >=1.3.x.x'),
              mterror, [mbok], 0);

            DB1.Disconnect;
            exit;
          END;
      EXCEPT
        //           V :=-1; // zur Sicherheit, damit die nä. Updateschritte nicht ausgeführt werden !!!

        MessageDlg(_('Die Datenbankversion konnte nicht ermittelt werden !' +
          #13#10 +
          'Der ausgewählte Mandant wurde nicht geöffnet.'), mterror, [mbok], 0);
        DB1.Disconnect;
        exit;
      END;

      //---ENDE Check----

      // DB Update-Check für Tabellenstruktur Version 1.09

      TRY
        S := ReadString('MAIN', 'DB_VERSION', '0.00');
        IF DecimalSeparator <> '.' THEN
          WHILE (Pos('.', S) > 0) DO
            S[Pos('.', S)] := DecimalSeparator;
        V := StrToFloat(S) * 100;
      EXCEPT
        //           V :=-1; // zur Sicherheit, damit die nä. Updateschritte nicht ausgeführt werden !!!

        MessageDlg('Die Datenbankversion konnte nicht ermittelt werden !' +
          #13#10 +
          'Der ausgewählte Mandant wurde nicht geöffnet.', mterror, [mbok], 0);
        DB1.Disconnect;
        exit;
      END;

      IF (V = 109) THEN
        BEGIN
          MessageDlg('Die ausgewählte Datenbank ist nicht auf dem aktuellen Stand'
            + #13#10 +
            'und wird deshalb nun aktualisiert (Version 1.09 auf Version 1.10 bzw. 1.11)',
            mtinformation, [mbok], 0);

          Res := GetDBUserRechte(True, '', '');
          IF (urSelect IN Res) AND (urInsert IN Res) AND (urUpdate IN Res) AND
            (urDelete IN Res) AND (urCreate IN Res) AND (urAlter IN Res) AND
            (urDrop IN Res) AND (urIndex IN Res) THEN
            BEGIN
              // User hat die benötigten Rechte !
              Application.ProcessMessages;

              IF NOT UpdateDataBase(DBUpdTo1_10, Warn, Error, 'db_update_110')
                THEN
                MessageDlg('Fehler beim Update der Tabellenstruktur auf Version 1.10 bzw. 1.11 !',
                  mterror, [mbok], 0);
            END
          ELSE
            BEGIN
              MsgNoSQLRights(Res);
              DB1.Disconnect;
              exit;
            END;
        END; //if

      IF (V = 110) THEN
        BEGIN
          MessageDlg('Die ausgewählte Datenbank ist nicht auf dem aktuellen Stand'
            + #13#10 +
            'und wird deshalb nun aktualisiert (Version 1.10 auf Version 1.11 )',
            mtinformation, [mbok], 0);

          Res := GetDBUserRechte(True, '', '');
          IF (urSelect IN Res) AND (urInsert IN Res) AND (urUpdate IN Res) AND
            (urDelete IN Res) AND (urCreate IN Res) AND (urAlter IN Res) AND
            (urDrop IN Res) AND (urIndex IN Res) THEN
            BEGIN
              // User hat die benötigten Rechte !
              Application.ProcessMessages;

              IF NOT UpdateDataBase(DBUpdTo1_11, Warn, Error, 'db_update_110')
                THEN
                MessageDlg('Fehler beim Update der Tabellenstruktur auf Version  1.11 !',
                  mterror, [mbok], 0);
            END
          ELSE
            BEGIN
              MsgNoSQLRights(Res);
              DB1.Disconnect;
              exit;
            END;
        END; //if

      IF (V = 111) THEN
        BEGIN
          MessageDlg('Die ausgewählte Datenbank ist nicht auf dem aktuellen Stand'
            + #13#10 +
            'und wird deshalb nun aktualisiert (Version 1.11 auf Version 1.12 )',
            mtinformation, [mbok], 0);

          Res := GetDBUserRechte(True, '', '');
          IF (urSelect IN Res) AND (urInsert IN Res) AND (urUpdate IN Res) AND
            (urDelete IN Res) AND (urCreate IN Res) AND (urAlter IN Res) AND
            (urDrop IN Res) AND (urIndex IN Res) THEN
            BEGIN
              // User hat die benötigten Rechte !
              Application.ProcessMessages;

              IF NOT UpdateDataBase(DBUpdTo1_11, Warn, Error, 'db_update_111')
                THEN
                MessageDlg('Fehler beim Update der Tabellenstruktur auf Version  1.12 !',
                  mterror, [mbok], 0);
            END
          ELSE
            BEGIN
              MsgNoSQLRights(Res);
              DB1.Disconnect;
              exit;
            END;
        END; //if

      IF (V = 112) THEN
        BEGIN
          MessageDlg('Die ausgewählte Datenbank ist nicht auf dem aktuellen Stand'
            + #13#10 +
            'und wird deshalb nun aktualisiert (Version 1.12 auf Version 1.13 )',
            mtinformation, [mbok], 0);

          Res := GetDBUserRechte(True, '', '');
          IF (urSelect IN Res) AND (urInsert IN Res) AND (urUpdate IN Res) AND
            (urDelete IN Res) AND (urCreate IN Res) AND (urAlter IN Res) AND
            (urDrop IN Res) AND (urIndex IN Res) THEN
            BEGIN
              // User hat die benötigten Rechte !
              Application.ProcessMessages;

              IF NOT UpdateDataBase(DBUpdTo1_13, Warn, Error, 'db_update_113')
                THEN
                MessageDlg('Fehler beim Update der Tabellenstruktur auf Version  1.13 !', mterror, [mbok], 0);
            END
          ELSE
            BEGIN
              MsgNoSQLRights(Res);
              DB1.Disconnect;
              exit;
            END;
        END; //if

      // DB Tabellenstruktur zu neu für das Programm ???
      TRY
        S := ReadString('MAIN', 'DB_VERSION', '0.00');
        IF DecimalSeparator <> '.' THEN
          WHILE (Pos('.', S) > 0) DO
            S[Pos('.', S)] := DecimalSeparator;
        V := StrToFloat(S) * 100;
      EXCEPT
        //           V :=-1; // zur Sicherheit, damit die nä. Updateschritte nicht ausgeführt werden !!!

        MessageDlg(_('Die Datenbankversion konnte nicht ermittelt werden !' + #13#10 +
          'Der ausgewählte Mandant wurde nicht geöffnet.'), mterror, [mbok], 0);
        DB1.Disconnect;
        exit;
      END;

      IF V = DBVersion_Soll THEN
        BEGIN

          FOR i := 0 TO DB1.DatasetCount - 1 DO
            IF DSTab[i] THEN
              tDataset(DB1.Datasets[i]).Active := True;

          setlength(DSTab, 0);

          AktMandant := NewMandant;

          Result := True;

          IF Save THEN
            BEGIN
              ininame := extractfilepath(paramstr(0)) + CAO32_DB_CFG;
              MyIni := tIniFile.Create(IniName);
              TRY
                MyIni.WriteString(APP, 'AKTMANDANT', AktMandant);
              FINALLY
                MyIni.Free;
              END;
            END;

          MandantOK := True;
        END
      ELSE
        BEGIN
          IF V > DBVersion_Soll THEN
            BEGIN
              MessageDlg(_('Der aktuelle Mandant hat eine neuere DB-Version' + #13#10 +
                'und kann mit dieser Programmversion nicht geöffnet werden.' + #13#10 +
                'Bitte aktualisieren Sie das Programm ...'), mterror, [mbok], 0);

              DB1.Disconnect;
              Result := False;
            END;
        END;
    END
  ELSE
    BEGIN
      // Mandantendaten konnten aus der INI nicht gelesen werden.
      Result := False;

      MessageDlg(_('Die Einstellungen für den aktuellen Mandanten' + #13#10 +
        'konnten aus der Datei "' + CAO32_DB_CFG + '" nicht gelesen werden.' + #13#10 +
        'Bitte prüfen Sie die Einstellungen.'), mterror, [mbok], 0);
    END;
  FirmaTab.Open;
  MitarbeiterTab.Open;
  MitarbeiterAktivTab.Open;
  ZweigTab.Open;
END;
//------------------------------------------------------------------------------
// Sollte nach dem Öffnen eines neuen Mandanten aufgerufen werden
//------------------------------------------------------------------------------

PROCEDURE TDM1.InitMandantAfterOpen;
VAR
  S : STRING;
  ini : tinifile;
  i, v : Integer;
BEGIN
  TRY
    WhrungTab.Open;
    LandTab.Open;
    LiefArtTab.Open;
    ZahlArtTab.Open;
    ShopOrderStatusTab.Open;
    VertreterTab.Open;
    ObjGruppenTab.Open;
    SprachTab.Open;

    IF SprachTab.RecordCount > 0 THEN
      BEGIN
        DefSpracheID := SprachTab.FieldByName('SPRACH_ID').AsInteger;
        DefSprachCode := SprachTab.FieldByName('CODE').AsString;
      END;

    // PLZ installieren
    TRY
      // Alte Version einlesen
      V := ReadInteger('MAIN', 'PLZ_VERSION', 106);

      uniquery.Close;
      uniquery.sql.text := 'select count(*) as ANZ from PLZ';
      uniquery.Open;
      IF (uniquery.recordcount = 1) AND
        (
        (uniquery.fieldbyname('ANZ').asInteger < 10) OR
        (PLZ_VERSION > V)
        ) AND
        (fileexists(extractfilepath(paramstr(0)) + 'plz.cao')) THEN
        BEGIN
          ProgressForm.Init(_('PLZ importieren...'));
          uniquery.close;
          uniquery.sql.text := 'delete from PLZ';
          uniquery.execsql;
          ZBatchSql1.Sql.LoadFromFile(extractfilepath(paramstr(0)) + 'plz.cao');
          TRY
            SQLLog := False;
            Screen.Cursor := crSqlWait;
            ZBatchSql1.ExecSql;
          FINALLY
            SQLLog := True;
            Screen.Cursor := crDefault;
          END;
          WriteInteger('MAIN', 'PLZ_VERSION', PLZ_VERSION);
        END;
      uniquery.close;
    EXCEPT
    END;
    ProgressForm.Stop;

    // BLZ installieren
    TRY
      // Alte Version einlesen
      V := ReadInteger('MAIN', 'BLZ_VERSION', 106);

      uniquery.Close;
      uniquery.sql.text := 'select count(*) as ANZ from BLZ';
      uniquery.Open;
      IF (uniquery.recordcount = 1) AND
        (
        (uniquery.fieldbyname('ANZ').asInteger < 10) OR
        (BLZ_VERSION > V)
        ) AND
        (fileexists(extractfilepath(paramstr(0)) + 'blz.cao')) THEN
        BEGIN
          ProgressForm.Init(_('BLZ importieren...'));
          uniquery.close;
          uniquery.sql.text := 'delete from BLZ';
          uniquery.execsql;
          ZBatchSql1.Sql.LoadFromFile(extractfilepath(paramstr(0)) + 'blz.cao');
          TRY
            SQLLog := False;
            Screen.Cursor := crSqlWait;
            ZBatchSql1.ExecSql;
          FINALLY
            SQLLog := True;
            Screen.Cursor := crDefault;
          END;
          WriteInteger('MAIN', 'BLZ_VERSION', BLZ_VERSION);
        END;
      uniquery.close;
    EXCEPT
    END;
    ProgressForm.Stop;

    // Länder installieren
    TRY
      uniquery.Close;
      uniquery.sql.text := 'select count(*) as ANZ from LAND';
      uniquery.Open;
      IF (uniquery.recordcount = 1) AND
        (uniquery.fieldbyname('ANZ').asInteger < 20) AND
        (fileexists(extractfilepath(paramstr(0)) + 'land.cao')) THEN
        BEGIN
          ProgressForm.Init(_('Länder importieren...'));
          ZBatchSql1.Sql.LoadFromFile(extractfilepath(paramstr(0)) +
            'land.cao');
          IF uniquery.fieldbyname('ANZ').asInteger > 0 THEN
            ZBatchSql1.Sql.Insert(0, 'delete from LAND;');
          ZBatchSql1.ExecSql;
        END;
      uniquery.close;
    EXCEPT
    END;
    ProgressForm.Stop;

    LeitWaehrung := ReadString('MAIN', 'LEITWAEHRUNG', '@@');
    IF LeitWaehrung = '@@' THEN
      BEGIN
        LeitWaehrung := '€';
        WriteString('MAIN', 'LEITWAEHRUNG', '€');
      END;

    LandK2 := ReadString('MAIN', 'LAND', 'DE');

    AnzPreis := ReadInteger('MAIN\ARTIKEL', 'ANZPREIS', -1);
    IF AnzPreis = -1 THEN
      BEGIN
        WriteInteger('MAIN\ARTIKEL', 'ANZPREIS', 5);
        AnzPreis := 5;
      END;

    EK_NACHKOMMA := ReadInteger('MAIN\BELEGE', 'EK_EP_NACHKOMMASTELLEN', 3);
    VK_NACHKOMMA := ReadInteger('MAIN\BELEGE', 'VK_EP_NACHKOMMASTELLEN', 2);

    IF VK_NACHKOMMA > 4 THEN
      VK_NACHKOMMA := 4;
    IF VK_NACHKOMMA < 2 THEN
      VK_NACHKOMMA := 2;

    IF EK_NACHKOMMA > 4 THEN
      EK_NACHKOMMA := 4;
    IF EK_NACHKOMMA < 2 THEN
      EK_NACHKOMMA := 2;

    EK_DFormat := ',#0.';
    FOR i := 1 TO EK_NACHKOMMA DO
      EK_DFormat := EK_DFormat + '0';
    EK_EFormat := '0.';
    FOR i := 1 TO EK_NACHKOMMA DO
      EK_EFormat := EK_EFormat + '0';

    VK_DFormat := ',#0.';
    FOR i := 1 TO VK_NACHKOMMA DO
      VK_DFormat := VK_DFormat + '0';
    VK_EFormat := '0.';
    FOR i := 1 TO VK_NACHKOMMA DO
      VK_EFormat := VK_EFormat + '0';

    // Globale MWST-Tabelle
    MWSTTab[0] := ReadDouble('MAIN\MWST', '0', 0);
    MWSTTab[1] := ReadDouble('MAIN\MWST', '1', 19); // 19% MWSt. 2007
    MWSTTab[2] := ReadDouble('MAIN\MWST', '2', 7);
    MWSTTab[3] := ReadDouble('MAIN\MWST', '3', 0);

    // Default-Steuer und Code laden
    I := ReadInteger('MAIN\MWST', 'DEFAULT', 1);
    IF (i < 0) OR (i > 3) THEN
      I := 2;
    DefMwStCD := I;
    DefMwSt := MWSTTab[I];

    // Globale Kalkulationsfaktoren
    FOR i := 1 TO 5 DO
      BEGIN
        GCalcFaktorTab[i] :=
          ReadDouble('MAIN\ARTIKEL', 'VK' + IntToStr(i) + '_CALC_FAKTOR', 0)
      END;

    // Shop-Calc-Faktor
    GCalcFaktorTab[6] := ReadDouble('MAIN\ARTIKEL', 'SHOP_CALC_FAKTOR', 0);

    // Brutto-Rundungswert für Artikel
    BR_RUND_WERT := ReadInteger('MAIN\ARTIKEL', 'BRUTTO_RUNDUNG_WERT', 0);
    IF BR_RUND_WERT < 1 THEN
      BR_RUND_WERT := 1; //mind 1 cent

    // Brutto-Summen-Rundungswert für Belege
    BR_SUM_RUND_WERT := ReadInteger('MAIN\BELEGE', 'BRUTTO_RUNDUNG_WERT', 0);
    IF BR_SUM_RUND_WERT < 1 THEN
      BEGIN
        BR_SUM_RUND_WERT := 1; //mind 1 cent
        WriteInteger('MAIN\BELEGE', 'BRUTTO_RUNDUNG_WERT', BR_SUM_RUND_WERT);
      END;

    i := ReadInteger('MAIN', 'USE_KFZ', -1);
    IF i = -1 THEN
      BEGIN
        i := 0;
        WriteInteger('MAIN', 'USE_KFZ', 0);
      END;

    Use_KFZ := i = 1;

    IF USE_KFZ THEN
      BEGIN
        S := ReadString('MAIN\KFZ', 'ATRIS_PFAD', '@@@');
        IF S = '@@@' THEN
          WriteString('MAIN\KFZ', 'ATRIS_PFAD', '');

        IF (length(S) > 0) AND (S <> '@@@') AND (DirectoryExists(S)) THEN
          BEGIN
            IF (length(s) > 0) AND (s[length(s)] <> '\') THEN
              s := s + '\';
            AtrisPfad := S;
            AtrisEnable := True;
          END
        ELSE
          BEGIN
            ini := tinifile.create('atris_st.ini');
            TRY
              s := ini.ReadString('INTERFACE', 'BESTINFO32', '');
            FINALLY
              ini.free;
            END;
            IF length(s) > 0 THEN
              BEGIN
                s := extractfilepath(s);
                IF (length(s) > 0) AND (s[length(s)] <> '\') THEN
                  s := s + '\';
                IF DirectoryExists(S) THEN
                  BEGIN
                    AtrisPfad := s;
                    AtrisEnable := True;
                  END;
              END;
          END;
      END;

    CheckTerminalID; // Terminal-Nummer auslesen und ggf. neu setzen

    // Default-Sprache laden
    SprachTab.Open;
    IF SprachTab.RecordCount > 0 THEN
      BEGIN
        DefSpracheID := SprachTab.FieldByName('SPRACH_ID').AsInteger;
        DefSprachCode := SprachTab.FieldByName('CODE').AsString;
      END;

    BLZ_LEN := ReadInteger('MAIN\ADRESSEN', 'BLZ_LEN', -1);
    IF BLZ_LEN = -1 THEN
      BEGIN
        BLZ_LEN := 8;
        WriteInteger('MAIN\ADRESSEN', 'BLZ_LEN', 8);
      END;

    // Pfade laden
    BackupDir := ReadString('MAIN\PFADE', 'BACKUP_DIR', MainDir + 'BACKUP\');
    TmpDir := ReadString('MAIN\PFADE', 'TMP_DIR', MainDir + 'TMP\');
    DTADir := ReadString('MAIN\PFADE', 'DTA_DIR', MainDir + 'DTA\');
    ExportDir := ReadString('MAIN\PFADE', 'EXPORT_DIR', MainDir + 'EXPORT\');
    ImportDir := ReadString('MAIN\PFADE', 'IMPORT_DIR', MainDir + 'IMPORT\');
    DokumentenDir := ReadString('MAIN\PFADE', 'DOKUMENTEN_DIR', MainDir + 'DOKUMENTEN\');

    IF (length(BackupDir) > 0) AND (Backupdir[length(BackupDir)] <> '\') THEN
      BackupDir := BackupDir + '\';
    IF (length(TmpDir) > 0) AND (TmpDir[length(TmpDir)] <> '\') THEN
      TmpDir := TmpDir + '\';
    IF (length(DTADir) > 0) AND (DTADir[length(DTADir)] <> '\') THEN
      DTADir := DTADir + '\';
    IF (length(ExportDir) > 0) AND (ExportDir[length(ExportDir)] <> '\') THEN
      ExportDir := ExportDir + '\';
    IF (length(ImportDir) > 0) AND (ImportDir[length(ImportDir)] <> '\') THEN
      ImportDir := ImportDir + '\';
    IF (length(DokumentenDir) > 0) AND (DokumentenDir[length(DokumentenDir)] <> '\') THEN
      DokumentenDir := DokumentenDir + '\';

    // wenn Pfad nicht existiert dann anlegen
    TRY
      ForceDirectories(BackupDir);
    EXCEPT
    END;
    TRY
      ForceDirectories(TmpDir);
    EXCEPT
    END;
    TRY
      ForceDirectories(DTADir);
    EXCEPT
    END;
    TRY
      ForceDirectories(ExportDir);
    EXCEPT
    END;
    TRY
      ForceDirectories(ImportDir);
    EXCEPT
    END;
    TRY
      ForceDirectories(DokumentenDir);
    EXCEPT
    END;
  EXCEPT
    MessageDlg(_('Beim öffnen der Tabellen ist ein Fehler aufgetreten !'),
      mterror, [mbok], 0);
  END;
END;

//------------------------------------------------------------------------------
// prüft mittels TryAndError die DB-Rechte des Benutzers ...
// zurückgeliefert wird ein Set der Rechte
// wenn es dumm läuft und der User zwar CREATE-Rechte hat, aber keine DROP-Rechte,
// dann bleibt die Tabelle UTEST als Leiche zurück
//------------------------------------------------------------------------------

FUNCTION TDM1.GetDBUserRechte(AktUser : Boolean; User, Secret : STRING) : tSDBUserRechte;
VAR
  LastUser,
    LastUserSecret : STRING;
  Error : Boolean; //S : String;
BEGIN
  Result := [];
  Error := False;
  IF AktUser = False THEN
    BEGIN
      // akt. User sichern
      LastUser := DB1.Login;
      LastUserSecret := DB1.Password;

      DB1.Disconnect;
      DB1.Login := User;
      DB1.Password := Secret;

      TRY
        DB1.Connect;
      EXCEPT
        Error := True;
      END;
    END;

  IF NOT Error THEN
    BEGIN
      // jetzt Rechte prüfen

      UniQuery.Close;
      UniQuery.Sql.Text := 'DROP TABLE IF EXISTS UCHECK ';
      TRY
        UniQuery.ExecSql;
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'CREATE TABLE UCHECK (SPALTE1 VARCHAR(10) ' +
        'NOT NULL, PRIMARY KEY(SPALTE1))';
      TRY
        UniQuery.ExecSql;
        include(Result, urCreate);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'ALTER TABLE UCHECK CHANGE SPALTE1 SPALTE1 ' +
        'INT(10) DEFAULT "1" NOT NULL';
      TRY
        UniQuery.ExecSql;
        include(Result, urAlter);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'CREATE INDEX TEST ON UCHECK (SPALTE1)';
      TRY
        UniQuery.ExecSql;
        include(Result, urIndex);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'INSERT INTO UCHECK (SPALTE1) VALUES (5)';
      TRY
        UniQuery.ExecSql;
        include(Result, urInsert);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'UPDATE UCHECK SET SPALTE1=6 WHERE SPALTE1=5';
      TRY
        UniQuery.ExecSql;
        include(Result, urUpdate);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'SELECT SPALTE1 FROM UCHECK WHERE SPALTE1=6';
      TRY
        UniQuery.Open;
        IF (UniQuery.Active) AND
          (UniQuery.RecordCount = 1) THEN
          include(Result, urSelect);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'DELETE FROM UCHECK WHERE SPALTE1=6';
      TRY
        UniQuery.ExecSql;
        include(Result, urDelete);
      EXCEPT

      END;
      UniQuery.Close;

      UniQuery.Sql.Text := 'DROP TABLE UCHECK';
      TRY
        UniQuery.ExecSql;
        include(Result, urDrop);
      EXCEPT

      END;
      UniQuery.Close;
    END;

  IF AktUser = False THEN
    BEGIN
      DB1.Disconnect;
      DB1.Login := LastUser;
      DB1.Password := LastUserSecret;

      TRY
        DB1.Connect;
      EXCEPT
        //           Error :=True;
      END;
    END;

  { // nur zu Testzwecken
  S :='';
  if urSelect in Result then S :=S+'SELECT : JA'+#13 else S :=S+'SELECT : NEIN'+#13;
  if urInsert in Result then S :=S+'INSERT : JA'+#13 else S :=S+'INSERT : NEIN'+#13;
  if urUpdate in Result then S :=S+'UPDATE : JA'+#13 else S :=S+'UPDATE : NEIN'+#13;
  if urDelete in Result then S :=S+'DELETE : JA'+#13 else S :=S+'DELETE : NEIN'+#13;
  if urCreate in Result then S :=S+'CREATE : JA'+#13 else S :=S+'CREATE : NEIN'+#13;
  if urAlter  in Result then S :=S+'ALTER : JA'+#13 else S :=S+'ALTER : NEIN'+#13;
  if urDrop   in Result then S :=S+'DROP : JA'+#13 else S :=S+'DROP : NEIN'+#13;
  //if urIndex  in Result then S :=S+'INDEX : JA'+#13 else S :=S+'INDEX : NEIN'+#13;

  //MessageDlg ('Userrechte :'+#13#10+S,mtinformation, [mbok],0);

  if (urSelect in Result)and
     (urInsert in Result)and
     (urUpdate in Result)and
     (urDelete in Result)and
     (urCreate in Result)and
     (urAlter in Result)and
     (urDrop in Result) then
  begin
    MessageDlg ('Benutzer darf Update ausführen !',mtinformation,[mbok],0);
  end; }
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.UpdateDatabase(Data : tJvStrHolder;
  VAR Warnings, Errors : Integer;
  LogFileName : STRING) : boolean;

VAR
  Idx : Integer;
  S, S1, S2 : STRING;
  F : TextFile;
  FN : STRING;
BEGIN
  FN := LogDir + LogFileName + '_' + FormatDateTime('yyyy_mm_dd_hh_mm_ss', Now) +
    '.log';

  AssignFile(F, FN);
  IF NOT fileexists(FN) THEN
    FileClose(FileCreate(FN));
  Append(F);

  Application.MainForm.Caption := '';

  Warnings := 0;
  Errors := 0;
  TRY
    UniQuery.Close;
    UniQuery.Sql.Clear;
    S2 := '';

    FOR idx := 0 TO Data.Strings.Count - 1 DO
      BEGIN
        s := Data.Strings[idx];
        S1 := s;
        WHILE (length(S1) > 0) AND (S1[length(S1)] = ' ') DO
          delete(S1, length(S1), 1);
        IF (pos(';', s1) > 0) AND (pos(';', s1) = length(s1)) THEN
          BEGIN
            delete(s, length(s1), 1);
            IF length(S) > 0 THEN
              S2 := S2 + S;
            TRY
              IF (length(S) > 0) AND (S[1] = ';') THEN
                delete(s, 1, 1);
              UniQuery.Sql.Text := S2;
              Application.MainForm.Caption := S2;
              UniQuery.ExecSql;
            EXCEPT
              ON e : exception DO
                BEGIN
                  inc(Warnings);

                  Writeln(F, 'SQL:' + S2);
                  Writeln(F, 'RES:' + e.Message);
                  Writeln(F);
                END;
            END;

            UniQuery.Close;
            UniQuery.Sql.Clear;
            S2 := '';
          END
        ELSE
          IF (length(S) > 0) AND (s[1] <> '#') THEN
            S2 := S2 + #13#10 + S;
      END;
    Result := True;
  EXCEPT
    inc(Errors);
    Result := False;
    UniQuery.Close;
    UniQuery.Sql.Clear;
  END;
  CloseFile(F);
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.IncNummer(Quelle : Integer) : Int64;
VAR
  F : Integer;
BEGIN
  Result := -1;
  IF (Quelle > 10) AND (Quelle < 20) THEN
    Quelle := 10;
  NummerTab.Close;
  NummerTab.Sql.Clear;
  NummerTab.Sql.Add('select VAL_INT as QUELLE, VAL_CHAR as FORMAT,');
  NummerTab.Sql.Add('VAL_INT2 as NEXT_NUM, VAL_INT3 as MAXLEN, MAINKEY, NAME');
  NummerTab.Sql.Add('from REGISTRY');
  NummerTab.Sql.Add('where MAINKEY="MAIN\\NUMBERS"');
  NummerTab.Sql.Add('and VAL_INT=:ID');

  NummerTab.ParamByName('ID').Value := Quelle;
  NummerTab.Open;
  IF NummerTab.RecordCount > 0 THEN
    BEGIN
      Result := NummerTabNext_Num.AsLargeInt;
      NummerTab.Edit;
      TRY
        F := Length(NummerTabFormat.AsString);
        NummerTabNext_Num.AsLargeInt := NummerTabNext_Num.AsLargeInt + 1;

        IF length(NummerTabNext_Num.AsString) > F THEN
          NummerTabNext_Num.AsLargeInt := 1;

        NummerTab.Post;
      EXCEPT
        NummerTab.Cancel;
      END;
    END
  ELSE
    BEGIN
      // Nummer existiert nicht
      InNewNummer := True;
      TRY
        NummerTab.Append;
        TRY
          NummerTabQUELLE.Value := Quelle;
          NummerTabNEXT_NUM.Value := 1;
          NummerTabFORMAT.Value := '000000';
          NummerTabMainKey.Value := 'MAIN\NUMBERS';
          NummerTabMAXLEN.AsInteger := 6;
          NummerTabNAME.AsString := IntToStr(QUELLE);
          NummerTab.Post;
          Result := 1;
        EXCEPT
          NummerTab.Cancel;
        END;
      FINALLY
        InNewNummer := False;
      END;
    END;
  NummerTab.Close;
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.IncNummerStr(Quelle : Integer) : STRING;
VAR
  Max : Integer;
  Num : Int64;
  Format : STRING;
BEGIN
  IF (Quelle > 10) AND (Quelle < 20) THEN
    Quelle := 10;
  NummerTab.Close;
  NummerTab.Sql.Clear;
  NummerTab.Sql.Add('select VAL_INT as QUELLE, VAL_CHAR as FORMAT,');
  NummerTab.Sql.Add('VAL_INT2 as NEXT_NUM, VAL_INT3 as MAXLEN, MAINKEY, NAME');
  NummerTab.Sql.Add('from REGISTRY');
  NummerTab.Sql.Add('where MAINKEY="MAIN\\NUMBERS"');
  NummerTab.Sql.Add('and VAL_INT=:ID');

  NummerTab.ParamByName('ID').Value := Quelle;
  NummerTab.Open;
  IF NummerTab.RecordCount > 0 THEN
    BEGIN
      Num := NummerTabNext_Num.AsLargeInt;
      NummerTab.Edit;
      TRY
        Max := NummerTabMaxLen.AsInteger;
        Format := NummerTabFormat.AsString;

        NummerTabNext_Num.AsLargeInt := Num + 1;

        IF Length(IntToStr(Num)) > Max THEN
          NummerTabNext_Num.AsLargeInt := 1;

        NummerTab.Post;
      EXCEPT
        NummerTab.Cancel;
      END;
    END
  ELSE
    BEGIN
      // Nummer existiert nicht
      Format := '000000';
      Num := 1;
      Max := 6;

      InNewNummer := True;
      TRY
        NummerTab.Append;
        TRY
          NummerTabQUELLE.Value := Quelle;
          NummerTabNEXT_NUM.Value := 1;
          NummerTabFORMAT.Value := Format;
          NummerTabMainKey.Value := 'MAIN\NUMBERS';
          NummerTabMAXLEN.AsInteger := Max;
          NummerTabNAME.AsString := IntToStr(QUELLE);
          NummerTab.Post;
        EXCEPT
          NummerTab.Cancel;
        END;
      FINALLY
        InNewNummer := False;
      END;
    END;
  NummerTab.Close;

  Result := FormatFloat(Format, Num);
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.GetNummerFormat(Quelle : Integer) : STRING;
BEGIN
  IF NummerTab.Active THEN
    NummerTab.Close;

  NummerTab.Sql.Clear;
  NummerTab.Sql.Add('select VAL_INT as QUELLE, VAL_CHAR as FORMAT,');
  NummerTab.Sql.Add('VAL_INT2 as NEXT_NUM, VAL_INT3 as MAXLEN, MAINKEY, NAME');
  NummerTab.Sql.Add('from REGISTRY');
  NummerTab.Sql.Add('where MAINKEY="MAIN\\NUMBERS"');
  NummerTab.Sql.Add('and VAL_INT=:ID');

  NummerTab.ParamByName('ID').Value := Quelle;
  NummerTab.Open;

  IF NummerTab.RecordCount = 0 THEN
    BEGIN
      InNewNummer := True;
      TRY
        NummerTab.Append;
        TRY
          NummerTabQUELLE.Value := Quelle;
          NummerTabNEXT_NUM.Value := 1;
          NummerTabFORMAT.Value := '000000';
          NummerTabMainKey.Value := 'MAIN\NUMBERS';
          NummerTabName.AsString := Inttostr(Quelle);
          NummerTab.Post;
        EXCEPT
          NummerTab.Cancel;
        END;
      FINALLY
        InNewNummer := False;
      END;
    END;

  Result := NummerTabFORMAT.AsString;

  NummerTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.ReadMandanten(App : STRING);
VAR
  ininame : STRING;
  ini : tinifile;
  po, idx, i : integer;
  S, S1, PW, PWC : STRING;

BEGIN
  idx := 0;
  ininame := extractfilepath(paramstr(0)) + CAO32_DB_CFG;
  ini := tinifile.create(ininame);
  //  DisplayTrace(ininame);
  TRY
    po := 1;
    WHILE (po < 1001) AND
      (ini.readstring('MANDANTEN',
      'M' + formatfloat('000', po),
      '@ERROR@') <> '@ERROR@') DO
      BEGIN
        idx := length(MandantTab);
        setlength(MandantTab, idx + 1);

        MandantTab[Idx].Name := ini.readstring('MANDANTEN', 'M' +
          formatfloat('000', po), '');
        MandantTab[Idx].Server := ini.readstring('MANDANTEN', 'M' +
          formatfloat('000', po) + '_SERVER', '');
        MandantTab[Idx].User := ini.readstring('MANDANTEN', 'M' +
          formatfloat('000', po) + '_USER', '');
        PW := ini.readstring('MANDANTEN', 'M' + formatfloat('000', po) + '_PASS',
          '');
        PWC := ini.readstring('MANDANTEN', 'M' + formatfloat('000', po) +
          '_PASS_C', '');

        DisplayTrace('User: ' + MandantTab[Idx].Name + ' ' +
          MandantTab[Idx].Server + ' ' + MandantTab[Idx].User);

        IF (length(PW) > 0) THEN
          BEGIN
            Cipher.Decoded := PW;
            PWC := Cipher.Encoded;
            IF length(PWC) > 0 THEN
              BEGIN
                S1 := '';
                FOR I := 1 TO length(PWC) DO
                  S1 := S1 + IntToHex(Ord(PWC[i]), 2);
                ini.writestring('MANDANTEN', 'M' + formatfloat('000', po) +
                  '_PASS_C', S1);
                ini.writestring('MANDANTEN', 'M' + formatfloat('000', po) +
                  '_PASS', '');
              END;
          END
        ELSE
          BEGIN
            IF length(PWC) >= 2 THEN
              BEGIN
                TRY
                  S1 := '';
                  FOR i := 1 TO length(PWC) DIV 2 DO
                    S1 := S1 + CHR(StrToInt('$' + Copy(PWC, (I - 1) * 2 + 1,
                      2)));
                  Cipher.Encoded := S1;
                  PW := Cipher.Decoded;
                EXCEPT
                  PW := '';
                END;
              END;
          END; //else if

        MandantTab[Idx].Pass := PW;
        MandantTab[Idx].DB := ini.readstring('MANDANTEN', 'M' +
          formatfloat('000', po) + '_DB', '');
        //NEU
        MandantTab[Idx].Port := ini.readinteger('MANDANTEN', 'M' +
          formatfloat('000', po) + '_PORT', 3306);
        //MandantTab[Idx].ShowLoginDlg  :=ini.ReadBool   ('MANDANTEN','M'+formatfloat ('000',po)+'_SHOW_LOGINDIALOG',False);
        //MandantTab[Idx].UseNTUserName :=ini.ReadBool   ('MANDANTEN','M'+formatfloat ('000',po)+'_USE_NTUSERNAME',False);

{$IFDEF AVE}
        MandantTab[Idx].UseSsh := ini.ReadBool('MANDANTEN', 'M' +
          formatfloat('000', po) + '_USE_SSH', False);
{$ENDIF}

        inc(po);
      END; //while

    AktMandant := Ini.ReadString(APP, 'AKTMANDANT', '');
    IF (AktMandant = '') AND (idx > 0) THEN
      AktMandant := MandantTab[0].Name;

    UseNTUserName := Ini.ReadBool(APP, 'USE_NTUSERNAME', True);
    DefaultUserName := Ini.ReadString(APP, 'DEFAULT_USER', '');
    DefaultPassword := Ini.ReadString(APP, 'DEFAULT_PASSWORD', '');
    PWC := Ini.ReadString(APP, 'DEFAULT_PASSWORD_C', '');

    //DisplayTrace(DefaultUserName);

    IF length(DefaultPassword) > 0 THEN
      BEGIN
        Cipher.Decoded := DefaultPassword;
        PWC := Cipher.Encoded;
        IF length(PWC) > 0 THEN
          BEGIN
            S1 := '';
            FOR I := 1 TO length(PWC) DO
              S1 := S1 + IntToHex(Ord(PWC[i]), 2);
            Ini.writestring(APP, 'DEFAULT_PASSWORD_C', S1);
            Ini.WriteString(APP, 'DEFAULT_PASSWORD', '');
          END;
      END
    ELSE
      IF length(PWC) > 2 THEN
        BEGIN
          TRY
            S1 := '';
            FOR i := 1 TO length(PWC) DIV 2 DO
              S1 := S1 + CHR(StrToInt('$' + Copy(PWC, (I - 1) * 2 + 1, 2)));
            Cipher.Encoded := S1;
            DefaultPassword := Cipher.Decoded;
          EXCEPT
            DefaultPassword := '';
          END;
        END; //if

    //PLZ- und BLZ-Verion lesen
    S := Ini.ReadString('VERSION', 'PLZ', '1.06');
    IF DecimalSeparator <> '.' THEN
      WHILE (Pos('.', S) > 0) DO
        S[Pos('.', S)] := DecimalSeparator;

    PLZ_VERSION := CAO_Round(StrToFloat(S) * 100);

    S := Ini.ReadString('VERSION', 'BLZ', '1.06');
    IF DecimalSeparator <> '.' THEN
      WHILE (Pos('.', S) > 0) DO
        S[Pos('.', S)] := DecimalSeparator;

    BLZ_VERSION := CAO_round(StrToFloat(S) * 100);

    //Kassen-Display DLL-Name ermitteln

    DisplayDLL := Ini.ReadString('DISPLAY', 'DLL_NAME', '');

  FINALLY
    ini.free;
  END;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.SaveMandanten;
VAR
  ininame : STRING;
  ini : tinifile;
  idx, i : integer;
  S1, PW, PWC : STRING;
  MySList : tStringList;

BEGIN
  IF length(MandantTab) = 0 THEN
    exit;

  ininame := extractfilepath(paramstr(0)) + CAO32_DB_CFG;
  ini := tinifile.create(ininame);
  MySList := tStringList.Create;
  TRY
    ini.ReadSectionValues('MANDANTEN', MySList);

    ini.EraseSection('MANDANTEN');

    FOR Idx := 0 TO length(MandantTab) - 1 DO
      BEGIN
        ini.Writestring('MANDANTEN', 'M' + formatfloat('000', Idx + 1),
          MandantTab[Idx].Name);
        ini.Writestring('MANDANTEN', 'M' + formatfloat('000', Idx + 1) +
          '_SERVER', MandantTab[Idx].Server);
        ini.Writestring('MANDANTEN', 'M' + formatfloat('000', Idx + 1) +
          '_USER', MandantTab[Idx].User);
        PW := MandantTab[Idx].Pass;
        IF (length(PW) > 0) THEN
          BEGIN
            Cipher.Decoded := PW;
            PWC := Cipher.Encoded;
            IF length(PWC) > 0 THEN
              BEGIN
                S1 := '';
                FOR I := 1 TO length(PWC) DO
                  S1 := S1 + IntToHex(Ord(PWC[i]), 2);
                ini.writestring('MANDANTEN', 'M' + formatfloat('000', IDX + 1) +
                  '_PASS_C', S1);
                ini.writestring('MANDANTEN', 'M' + formatfloat('000', IDX + 1) +
                  '_PASS', '');
              END;
          END;

        ini.Writestring('MANDANTEN', 'M' + formatfloat('000', Idx + 1) + '_DB',
          MandantTab[Idx].DB);
        //NEU
        ini.Writeinteger('MANDANTEN', 'M' + formatfloat('000', Idx + 1) +
          '_PORT', MandantTab[Idx].Port);
        ini.WriteBool('MANDANTEN', 'M' + formatfloat('000', Idx + 1) +
          '_SHOW_LOGINDIALOG', False);
        ini.WriteBool('MANDANTEN', 'M' + formatfloat('000', Idx + 1) +
          '_USE_NTUSERNAME', False);
      END;

    // aktuelle Mandanten der einzelnen programme zurückschreiben
    IF MySList.Count > 0 THEN
      BEGIN
        REPEAT
          IF Pos('CAO', Uppercase(MySList.Strings[0])) = 1 THEN
            BEGIN
              ini.Writestring('MANDANTEN',
                MySList.Names[0],
                MySList.Values[MySList.Names[0]]
                );
            END;
          MySList.Delete(0);

        UNTIL MySList.Count = 0;
      END;
  FINALLY
    ini.free;
    MySList.Free;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.GetMandant(Name : STRING; VAR Daten : MandantRec) : Boolean;
VAR
  po : integer;
BEGIN
  Result := False;
  name := uppercase(name);
  IF length(MandantTab) > 0 THEN
    FOR po := 0 TO length(mandanttab) - 1 DO
      BEGIN
        IF uppercase(mandanttab[po].name) = name THEN
          BEGIN
            Daten := mandanttab[po];
            Result := True;
            Break;
          END;
      END;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.NewMandant(Daten : MandantRec);
VAR
  I : Integer;
BEGIN
  I := length(MandantTab);
  SetLength(MandantTab, I + 1);
  MandantTab[i] := Daten;
  SaveMandanten;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.DeleteMandant(Name : STRING);
VAR
  I, J : Integer;
BEGIN
  IF length(MandantTab) = 0 THEN
    exit;
  FOR I := 0 TO length(MandantTab) - 1 DO
    BEGIN
      IF MandantTab[i].Name = Name THEN
        BEGIN
          IF I < length(MandantTab) - 1 THEN
            BEGIN
              FOR J := I + 1 TO length(MandantTab) - 1 DO
                BEGIN
                  MandantTab[J - 1] := MandantTab[j];
                END;
            END;
          SetLength(MandantTab, length(MandantTab) - 1);
        END;
    END;
  SaveMandanten;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadInteger(Key, Name : STRING; Default : Integer) : Integer;
BEGIN
  Result := -1;
  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF RegTab.RecordCount > 0 THEN
      Result := RegTabVal_Int.AsInteger
    ELSE
      Result := Default;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadInteger2(Key, Name : STRING; Default : Integer) : Integer;
BEGIN
  Result := -1;
  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF RegTab.RecordCount > 0 THEN
      Result := RegTabVal_Int2.AsInteger
    ELSE
      Result := Default;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;

//------------------------------------------------------------------------------

FUNCTION tDM1.ReadIntegerU(Key, Name : STRING; Default : Integer) : Integer;
BEGIN
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;

  Result := ReadInteger(Key, Name, Default);
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteInteger(Key, Name : STRING; Value : Integer);
BEGIN
    If not (value > (-5)) then
    begin
      showmessage('Write integer kleiner -5');
    end;
  IF NOT DB1.Connected THEN
    exit;

  RegTab.Close;
  RegTab.ParamByName('KEY').AsString := Key;
  RegTab.ParamByName('NAME').AsString := Name;
  RegTab.Open;
  IF RegTab.RecordCount = 0 THEN
    BEGIN
      RegTab.Append;
      TRY
        RegTabMainKey.Value := Key;
        RegTabName.Value := Name;
        RegTabVal_Int.AsInteger := Value;
        RegTabVal_Typ.AsInteger := 3;
        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END
  ELSE
    IF RegTabVal_Int.AsInteger <> Value THEN
      BEGIN
        RegTab.Edit;
        TRY
          RegTabVal_Int.AsInteger := Value;
          RegTabVal_Typ.AsInteger := 3;
          RegTab.Post;
        EXCEPT
          RegTab.Cancel;
        END;
      END;

  RegTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteIntegerU(Key, Name : STRING; Value : Integer);
BEGIN
  If not (value > (-5)) then
    begin
      showmessage('Writeinteger u kleiner -5');
    end;
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  WriteInteger(Key, Name, Value);
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadBoolean(Key, Name : STRING; Default : Boolean) : Boolean;
BEGIN
  Result := FALSE;

  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF RegTab.RecordCount > 0 THEN
      Result := RegTabVal_Int.AsInteger = 1
    ELSE
      Result := Default;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadBooleanU(Key, Name : STRING; Default : Boolean) : Boolean;
BEGIN
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;

  Result := ReadBoolean(Key, Name, Default);
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteBoolean(Key, Name : STRING; Value : Boolean);
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.Close;
  RegTab.ParamByName('KEY').AsString := Key;
  RegTab.ParamByName('NAME').AsString := Name;
  RegTab.Open;
  IF RegTab.RecordCount = 0 THEN
    BEGIN
      RegTab.Append;
      TRY
        RegTabMainKey.Value := Key;
        RegTabName.Value := Name;
        RegTabVal_Int.AsInteger := ord(Value);
        RegTabVal_Typ.AsInteger := 3;
        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END
  ELSE
    IF RegTabVal_Int.AsInteger <> ord(Value) THEN
      BEGIN
        RegTab.Edit;
        TRY
          RegTabVal_Int.AsInteger := ord(Value);
          RegTabVal_Typ.AsInteger := 3;
          RegTab.Post;
        EXCEPT
          RegTab.Cancel;
        END;
      END;

  RegTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteBooleanU(Key, Name : STRING; Value : Boolean);
BEGIN
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;

  WriteBoolean(Key, Name, Value);
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadString(Key, Name, Default : STRING) : STRING;
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF RegTab.RecordCount > 0 THEN
      Result := RegTabVal_Char.AsString
    ELSE
      Result := Default;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadStringU(Key, Name, Default : STRING) : STRING;
BEGIN
  Result := '';
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  Result := ReadString(Key, Name, Default);
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteString(Key, Name, Value : STRING);
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.Close;
  RegTab.ParamByName('KEY').AsString := Key;
  RegTab.ParamByName('NAME').AsString := Name;
  RegTab.Open;
  IF RegTab.RecordCount = 0 THEN
    BEGIN
      RegTab.Append;
      TRY
        RegTabMainKey.Value := Key;
        RegTabName.Value := Name;
        RegTabVal_Char.AsString := Value;
        RegTabVal_Typ.AsInteger := 1;
        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END
  ELSE
    IF RegTabVal_Char.AsString <> Value THEN
      BEGIN
        RegTab.Edit;
        TRY
          RegTabVal_Char.AsString := Value;
          RegTabVal_Typ.AsInteger := 1;
          RegTab.Post;
        EXCEPT
          RegTab.Cancel;
        END;
      END;

  RegTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteStringU(Key, Name, Value : STRING);
BEGIN
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  WriteString(Key, Name, Value);
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadLongString(Key, Name, Default : STRING) : STRING;
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF RegTab.RecordCount > 0 THEN
      Result := RegTabVal_Blob.AsString
    ELSE
      Result := Default;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadLongStringU(Key, Name, Default : STRING) : STRING;
BEGIN
  Result := '';
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  Result := ReadLongString(Key, Name, Default);
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteLongString(Key, Name, Value : STRING);
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.Close;
  RegTab.ParamByName('KEY').AsString := Key;
  RegTab.ParamByName('NAME').AsString := Name;
  RegTab.Open;
  IF RegTab.RecordCount = 0 THEN
    BEGIN
      RegTab.Append;
      TRY
        RegTabMainKey.Value := Key;
        RegTabName.Value := Name;
        RegTabVal_Blob.AsString := Value;
        RegTabVal_Typ.AsInteger := 5;
        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END
  ELSE
    IF RegTabVal_Blob.AsString <> Value THEN
      BEGIN
        RegTab.Edit;
        TRY
          RegTabVal_Blob.AsString := Value;
          RegTabVal_Typ.AsInteger := 5;
          RegTab.Post;
        EXCEPT
          RegTab.Cancel;
        END;
      END;

  RegTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteLongStringU(Key, Name, Value : STRING);
BEGIN
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  WriteLongString(Key, Name, Value);
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadDouble(Key, Name : STRING; Default : Double) : Double;
BEGIN
  Result := 0.0;
  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF RegTab.RecordCount > 0 THEN
      Result := RegTabVal_Double.AsFloat
    ELSE
      Result := Default;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadDoubleU(Key, Name : STRING; Default : Double) : Double;
BEGIN
  //     Result :=0;
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  Result := ReadDouble(Key, Name, Default);
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteDouble(Key, Name : STRING; Value : Double);
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.Close;
  RegTab.ParamByName('KEY').AsString := Key;
  RegTab.ParamByName('NAME').AsString := Name;
  RegTab.Open;
  IF RegTab.RecordCount = 0 THEN
    BEGIN
      RegTab.Append;
      TRY
        RegTabMainKey.Value := Key;
        RegTabName.Value := Name;
        RegTabVal_Double.AsFloat := Value;
        RegTabVal_Typ.AsInteger := 4;
        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END
  ELSE
    IF RegTabVal_Double.AsFloat <> Value THEN
      BEGIN
        RegTab.Edit;
        TRY
          RegTabVal_Double.AsFloat := Value;
          RegTabVal_Typ.AsInteger := 4;
          RegTab.Post;
        EXCEPT
          RegTab.Cancel;
        END;
      END;

  RegTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteDoubleU(Key, Name : STRING; Value : Double);
BEGIN
  IF length(Key) > 0 THEN
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER + '\' + KEY
  ELSE
    Key := 'USERSETTINGS\' + CAOSecurity.CurrUSER;
  WriteDouble(Key, Name, Value);
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.ReadLayout(Key, Name : STRING; VAR Data : tStream; Version :
  Integer = 0) : Boolean;
BEGIN
  Result := FALSE;
  IF NOT DB1.Connected THEN
    exit;

  RegTab.RequestLive := False;
  TRY
    RegTab.Close;
    RegTab.ParamByName('KEY').AsString := Key;
    RegTab.ParamByName('NAME').AsString := Name;
    RegTab.Open;
    IF (RegTab.RecordCount > 0) AND
      (NOT RegTabVal_Bin.IsNull) THEN
      BEGIN
        Data.Size := 0;
        Data.Position := 0;
        RegTabVal_Bin.SaveToStream(Data);
        Data.Position := 0;
        Result := RegTabVal_Int.AsInteger = Version;
      END;
    RegTab.Close;
  FINALLY
    RegTab.RequestLive := True;
  END;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.WriteLayout(Key, Name : STRING; Data : tStream; Version : Integer
  = 0);
BEGIN
  IF NOT DB1.Connected THEN
    exit;

  RegTab.Close;
  RegTab.ParamByName('KEY').AsString := Key;
  RegTab.ParamByName('NAME').AsString := Name;
  RegTab.Open;
  IF RegTab.RecordCount = 0 THEN
    BEGIN
      RegTab.Append;
      TRY
        RegTabMainKey.Value := Key;
        RegTabName.Value := Name;
        RegTabVAL_TYP.Value := 7;
        RegTabVal_Int.AsInteger := Version;
        Data.Position := 0;
        RegTabVAL_BIN.LoadFromStream(Data);

        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END
  ELSE
    BEGIN
      RegTab.Edit;
      TRY
        Data.Position := 0;
        RegTabVAL_BIN.LoadFromStream(Data);
        RegTabVAL_TYP.Value := 7;
        RegTabVal_INT.AsInteger := Version;
        RegTab.Post;
      EXCEPT
        RegTab.Cancel;
      END;
    END;

  RegTab.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE tDM1.GridSaveLayout(Grid : tDBGrid; Sec : STRING; Version : Integer =
  0);
VAR
  M : tMemoryStream;
BEGIN
  m := tmemorystream.create;
  TRY
    Grid.Columns.SaveToStream(M);
    WriteLayout('USERSETTINGS\' + CAOSecurity.CurrUser + '\LAYOUT', SEC, M,
      Version);
  FINALLY
    M.Free;
  END;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.GridLoadLayout(VAR Grid : tDBGrid; Sec : STRING; Version : Integer
  = 0);
VAR
  M : tMemoryStream;
  //I : Integer;
  //Found : Boolean;
BEGIN
  M := tMemoryStream.Create;
  TRY
    TRY
      IF ReadLayout('USERSETTINGS\' + CAOSecurity.CurrUser + '\LAYOUT', SEC,
        tStream(M), Version) THEN
        Grid.Columns.LoadFromStream(M);
      {
      if Grid.Columns.Count>0 then
      begin
        repeat
          Found :=False;
          for i:=0 to Grid.Columns.Count-1 do
          begin
            if not assigned(Grid.Columns[i].Field) then
            begin
              Grid.Columns[i].Free;
              Found :=True;
              Break;
            end;
          end;
        until not Found;
      end;
      }
    EXCEPT
    END;
  FINALLY
    M.Free;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Rechnung(Journal_ID : Integer) : STRING; // liefert Rechnungnummer zurück
VAR
  Pos : Integer;
  //IStr     : String;
  SNSql : STRING;
  gesamtrabatt : double;
BEGIN
  Result := '';

  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;

      // neue Rechnungsnummer holen, aber nur wenn USE_SHOP_ORDERID=False

      IF (NOT ReadBoolean('SHOP', 'USE_SHOP_ORDERID', False)) OR
        (JourTabSHOP_ID.AsInteger < 1) THEN
        JourTabVRENUM.Value := IncNummerStr(VK_RECH);

      Result := JourTabVRENUM.Value;

      JourTabSTADIUM.Value := 22;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := VK_RECH;
      JourTabQuelle_Sub.Value := 1; // Rechnung

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabFreigabe1_Flag.Value := False;
      JourTabMAHNKOSTEN.Value := 0;

      // Kundendaten (Zahlungsart und Lieferart) aktualisieren,
      // falls diese noch nicht zugewiesen sind
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;

      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKun_Zahlart.AsInteger < 0 THEN
            KunTabKun_Zahlart.Value := JourTabZahlart.Value;
          IF KunTabKun_Liefart.AsInteger < 0 THEN
            KunTabKun_Liefart.Value := JourTabLiefart.Value;
          KunTab.Post;
        END;

      KunTab.Close;

      // Automatikpositionen ( Gesamtrabatt, usw ) eintragen
      IF DM1.ReadBoolean('MAIN\Rabatt', 'Gesamtrabattzeile', false) THEN
        BEGIN
          UniQuery.Close;
          UniQuery.Sql.Text :=
            'select * from rabattgruppen where rabgrp_typ = 22 order by min_menge';
          UniQuery.Open;
          gesamtrabatt := 0;
          WHILE NOT Uniquery.eof DO
            BEGIN
              IF JourTabBsumme.Value >
                Uniquery.FieldByName('Min_Menge').AsInteger THEN
                gesamtrabatt := Uniquery.FieldByName('Rabatt2').Asfloat
              ELSE
                break;
              Uniquery.next;
            END;

          IF Gesamtrabatt > 0 THEN
            BEGIN
              TRY
                JPosTab.Close;
                JPosTab.Open;
                JPosTab.Append;
                JPosTabJournal_ID.Value := JourTabREC_ID.Value;
                JPosTabQuelle.Value := JourTabQuelle.Value;
                JPosTabQuelle_Sub.Value := JourTabQuelle_Sub.Value;
                JPosTabAddr_ID.Value := JourTabADDR_ID.Value;
                JPosTabVRENUM.Value := JourTabVRENUM.Value;
                JPosTabVLSNUM.Value := '';
                JPosTabATRNum.Value := -1;
                JPosTabBezeichnung.AsString := 'Abzüglich Gesamtrabatt ' +
                  floattostr(gesamtrabatt) + '%';
                JPosTabMenge.AsFloat := -1;
                JPosTabSTEUER_CODE.Value := 1;
                JPosTabEPreis.AsFloat := JourTabNsumme.AsFloat / 100 *
                  gesamtrabatt;
                JPosTabGPREIS.AsFloat := JPosTabEPreis.AsFloat * -1;
                JPosTabPOSITION.AsInteger := 999;
                JPosTab.Post;
              EXCEPT
                JPosTab.Cancel;
                Showmessage('Fehler beim Rabatt anfügen');
                result := '';
                exit;
              END;
              JourTabNSumme.AsFloat := JourTabNSumme.AsFloat +
                (JPosTabGPREIS.AsFloat);
              JourTabMSumme.AsFloat := JourTabMSumme.AsFloat +
                (JPosTabGPREIS.AsFloat / 100 * JourTabMWST_1.AsFloat);
              JourTabMSUMME_1.AsFloat := JourTabMSUMME_1.AsFloat +
                (JPosTabGPREIS.AsFloat / 100 * JourTabMWST_1.AsFloat);
              JourTabBsumme.AsFloat := JourTabBsumme.AsFloat +
                (JPosTabGPREIS.AsFloat + (JPosTabGPREIS.AsFloat / 100 *
                JourTabMWST_1.AsFloat));
            END; // if gesamtrabatt > 0

        END; // Ende If Gesamtrabatt

      // Ende -> Automatikpositionen ( Gesamtrabatt, usw ) eintragen
      CASE JourTabZahlart.Value OF
        //bar bzw. scheck
        1, 5 :
          BEGIN
            IF JourTabSOLL_SKONTO.Value > 0 THEN
              BEGIN
                JourTabStadium.Value := 80 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := JourTabSOLL_SKONTO.Value;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value :=
                  JourTabBSumme.Value - (JourTabBSumme.Value / 100) *
                  JourTabSOLL_SKONTO.Value;
              END
            ELSE
              BEGIN
                JourTabStadium.Value := 90 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := 0;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value := JourTabBSumme.Value;
              END;
          END;
        //Überweisung
        2, 3, 4 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
        //Lastschrift, EC-Karte
        6, 9 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
      END;

      // Wenn MWST-Freie Rechnung dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      IF JourTabZahlArt.Value = 1 THEN // Kassenbuchung
        BEGIN
          BucheKasse(JourTabIST_Zahldat.Value,
            VK_RECH, JourTabRec_ID.Value,
            JourTabVReNum.Value,
            JourTabGegenKonto.Value,
            JourTabIST_Skonto.Value,
            JourTabIST_Betrag.Value,
            'ZE VK-RE ' + JourTabKun_Name1.Value);

        END;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      //IStr :=''; // Insert-String für Stücklistenartikel leeren

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Artikel Buchen
          IF (JPosTabGebucht.Value = False) AND
            (
            (JPosTabArtikelTyp.Value = 'N') OR
            (JPosTabArtikelTyp.Value = 'X')
            ) AND
            (JPosTabARTIKEL_ID.Value > -1) THEN
            BEGIN
              // Menge erniedrigen
              ArtMengeTab.Close;
              ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
              ArtMengeTab.ParamByName('SUBMENGE').Value := JPosTabMenge.Value;
              ArtMengeTab.ExecSql;

              JPosTabGebucht.Value := True;
            END;
          {
             else
          if (JPosTabGebucht.Value=False)and
             (JPosTabArtikelTyp.Value='S')and
             (JPosTabARTIKEL_ID.Value>-1) then
          begin
               // Stückliste abarbeiten

               // gut währe es, die Stücklistenartikel versteckt mit in die Rechnung einzufügen,
               // dann würde man den verkauf auch in der Historie der Artikel sehen,
               // allerdings müßte es dann ein Flag Visible in der Postentabelle geben,
               // damit die Berechnungen der Preise und der Rechnungsausdruck korrekt funktionieren !!!

               STListTab.Close;
               STListTab.ParamByName ('ID').ASInteger :=JPosTabARTIKEL_ID.Value;
               STListTab.Open;

               while not STListTab.Eof do
               begin
                   // Menge erniedrigen
                   ArtMengeTab.Close;
                   ArtMengeTab.ParamByName ('ID').Value :=STListTabART_ID.Value;
                   ArtMengeTab.ParamByName ('SUBMENGE').Value :=JPosTabMenge.Value * STListTabMENGE.Value;
                   //Bestellmenge nicht verändern
                   //ArtMengeTab.ParamByName ('BMENGE').Value :=0;

                   ArtMengeTab.ExecSql;

                   if length(IStr)>0 then IStr :=IStr+';'+#13#10;

                   // Batch-SQL erzeugen und die Stücklistenartikel mit in die Rechnung zu speichern
                   // mit Artikeltyp="X"

                   dm1.uniquery.close;
                   dm1.uniquery.sql.text :='select MATCHCODE,ARTNUM,BARCODE,LAENGE,'+
                                           'GROESSE,DIMENSION,GEWICHT,ME_EINHEIT,'+
                                           'LANGNAME from ARTIKEL where REC_ID='+
                                           IntToStr(STListTabART_ID.Value);
                   dm1.uniquery.open;

                   IStr :=Istr+
                     'INSERT INTO JOURNALPOS SET '+
                     'QUELLE='+IntToStr(VK_RECH)+
                     ',QUELLE_SUB='+IntToStr(1)+
                     ',JOURNAL_ID='+IntToStr(JPosTabJOURNAL_ID.Value)+
                     ',ARTIKELTYP="X"'+
                     ',ARTIKEL_ID='+IntToStr(STListTabART_ID.Value)+
                     ',TOP_POS_ID='+IntToStr(JPosTabRec_ID.Value)+
                     ',ADDR_ID='+IntToStr(JPosTabADDR_ID.Value)+
                     ',VRENUM="'+JourTabVRENUM.Value+'"'+
                     ',MENGE="'+FloatToStrEx(JPosTabMenge.Value * STListTabMENGE.Value)+'"'+
                     ',POSITION='+IntToStr(JPosTabPOSITION.Value)+
                     ',MATCHCODE="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('MATCHCODE').AsString)+'"'+
                     ',ARTNUM="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('ARTNUM').AsString)+'"'+
                     ',BARCODE="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('BARCODE').AsString)+'"'+
                     ',LAENGE="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('LAENGE').AsString)+'"'+
                     ',GROESSE="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('GROESSE').AsString)+'"'+
                     ',DIMENSION="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('DIMENSION').AsString)+'"'+
                     ',GEWICHT='+FloatToStrEx(dm1.uniquery.fieldbyname ('GEWICHT').AsFloat)+
                     ',ME_EINHEIT="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('ME_EINHEIT').AsString)+'"'+
                     ',BEZEICHNUNG="'+ZSqlTypes.StringToSql(dm1.uniquery.fieldbyname ('LANGNAME').AsString)+'"';

                   dm1.uniquery.close;
                   STListTab.Next;
               end;
               STListTab.Close;

               JPosTabGebucht.Value :=True;
          end;   }

          // Daten aktualisieren
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := VK_RECH;
          JPosTabQuelle_Sub.Value := 1; // Rechnung

          JPosTab.Post;

          // Lieferschein(e) aktualisieren
          IF (JPosTabVLSNum.AsString <> '') AND
            (JPosTabVLSNum.AsString <> '0') THEN
            BEGIN
              uniquery.close;
              uniquery.sql.clear;
              uniquery.sql.add('update JOURNAL');
              uniquery.sql.add('set STADIUM=90, VRENUM="' + JPosTabVRENum.Value +
                '", RDATUM=:RDATUM');
              uniquery.sql.add('where QUELLE=2 and VLSNUM="' +
                JPosTabVLSNum.AsString + '"');
              uniquery.parambyname('RDATUM').AsDateTime := JourTabRDatum.Value;
              uniquery.ExecSql;
              uniquery.close;
              uniquery.sql.clear;
            END;
          JPosTab.Next;
        END;

      // im Artkikel die MENGE_VKRE_EDI aktualisieren
      UpdateArtikelEdiMenge(VK_RECH_EDI, JPosTabArtikel_ID.AsInteger, 0);

      {
      // Stücklistenunterartikel hinzufügen
      if length(IStr)>0 then
      begin
         DM1.ZBatchSql1.Sql.Text :=ISTr;
         try
            DM1.ZBatchSql1.ExecSql;
         except
            MessageDlg ('Fehler beim hinzufügen der Stücklisten-Unterartikel.',
                        mterror,[mbok],0);
         end;
      end;
      }
      IF JourTabKFZ_ID.Value >= 0 THEN
        BEGIN
          ReKFZTab.Close;
          ReKFZTab.ParamByName('KID').AsInteger := JourTabKFZ_ID.Value;
          ReKFZTab.Open;
          IF ReKFZTab.RecordCount = 1 THEN
            BEGIN
              ReKFZTab.Edit;
              ReKFZTabLE_BESUCH.Value := JourTabRDAtum.Value;
              ReKFZTabKM_STAND.Value := JourTabKM_Stand.Value;

              TRY
                ReKFZTab.Post;
              EXCEPT ReKFZTab.Cancel;
              END;
            END;
          ReKFZTab.Close;
        END;

      //Seriennumern aktualisieren
      TRY
        {
        dm1.UniQuery.close;
        dm1.UniQuery.sql.text :='UPDATE ARTIKEL_SERNUM SET VERK_NUM='+
                                IntToStr(JourTabVRENUM.AsInteger)+
                                ' where VK_JOURNAL_ID='+
                                IntToStr(JourTabRec_ID.AsInteger);
        dm1.UniQuery.ExecSql;
        dm1.UniQuery.close;  }

        //NEU
        UniQuery.Close;
        UniQuery.Sql.Text :=
          'select JPS.QUELLE,JPS.JOURNAL_ID,JPS.JOURNALPOS_ID,' +
          'JPS.ARTIKEL_ID,JPS.SNUM_ID from JOURNALPOS as JP, ' +
          'JOURNALPOS_SERNUM as JPS where JP.JOURNAL_ID=' +
          IntToStr(JourTabRec_ID.AsInteger) +
          ' and JP.SN_FLAG="Y" and JP.MENGE>0 and ' +
          'JP.ARTIKEL_ID=JPS.ARTIKEL_ID and ' +
          'JP.REC_ID=JPS.JOURNALPOS_ID and ' +
          'JP.JOURNAL_ID=JPS.JOURNAL_ID';

        SNSql := '';
        UniQuery.Open;
        WHILE NOT UniQuery.Eof DO
          BEGIN
            SNSql := SNSql +
              'UPDATE ARTIKEL_SERNUM SET STATUS="VK_RECH" ' +
              'WHERE ARTIKEL_ID=' +
              IntToStr(Uniquery.FieldByName('ARTIKEL_ID').AsInteger) +
              ' and SNUM_ID=' +
              IntToStr(Uniquery.FieldByName('SNUM_ID').AsInteger) + ';';
            UniQuery.Next;
          END;
        UniQuery.Close;

        ZBatchSql1.Sql.Text :=
          'UPDATE JOURNALPOS_SERNUM SET QUELLE=3 ' +
          'where JOURNAL_ID=' + IntToStr(JourTabRec_ID.AsInteger) + ';';

        IF length(SNSql) > 0 THEN
          ZBatchSql1.Sql.Add(SNSql);

        ProgressForm.Init(_('Seriennummern aktualisieren'));
        TRY
          ZBatchSql1.ExecSql;
        FINALLY
          ProgressForm.Stop;
        END;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der Seriennummern.'),
          mterror, [mbok], 0);
      END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen der Rechnung !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_BCKasse(Journal_ID : Integer) : STRING; // liefert BON-Nummer zurück
VAR
  Pos : Integer;
  SNSql : STRING;
BEGIN
  Result := '';
  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      // neue Rechnungsnummer holen
      JourTabVRENUM.Value := IncNummerStr({VK_KASSE}22);

      Result := JourTabVRENUM.Value;

      JourTabSTADIUM.Value := 22;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := VK_RECH;
      JourTabQuelle_Sub.Value := 2; // Kasse

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;
      JourTabFreigabe1_Flag.Value := False;

      CASE JourTabZahlart.Value OF
        //bar bzw. scheck oder ec-karte
        1, 5, 6 :
          BEGIN
            IF JourTabSOLL_SKONTO.Value > 0 THEN
              BEGIN
                JourTabStadium.Value := 80 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := JourTabSOLL_SKONTO.Value;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value :=
                  JourTabBSumme.Value - (JourTabBSumme.Value / 100) *
                  JourTabSOLL_SKONTO.Value;
              END
            ELSE
              BEGIN
                JourTabStadium.Value := 90 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := 0;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value := JourTabBSumme.Value;
              END;
          END;
      END;

      // Wenn MWST-Freier Kassenbeleg dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      {    if JourTabZahlArt.Value = 1 then  // Kassenbuchung
          begin
               BucheKasse (JourTabIST_Zahldat.Value,
                           VK_RECH,JourTabRec_ID.Value,
                           Inttostr(JourTabVReNum.Value),
                           JourTabGegenKonto.Value,
                           JourTabIST_Skonto.Value,
                           JourTabIST_Betrag.Value,
                           'ZE VK-RE '+JourTabKun_Name1.Value);

          end; }

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Artikel Buchen
          IF (JPosTabGebucht.Value = False) AND
            (
            (JPosTabArtikelTyp.Value = 'N') OR
            (JPosTabArtikelTyp.Value = 'X')
            ) AND
            (JPosTabARTIKEL_ID.Value > -1) THEN
            BEGIN
              // Menge erniedrigen
              ArtMengeTab.Close;
              ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
              ArtMengeTab.ParamByName('SUBMENGE').Value := JPosTabMenge.Value;
              ArtMengeTab.ExecSql;
              JPosTabGebucht.Value := True;
            END;
          {
             else
          if (JPosTabGebucht.Value=False)and
             (JPosTabArtikelTyp.Value='S')and
             (JPosTabARTIKEL_ID.Value>-1) then
          begin
             // Stückliste abarbeiten
             STListTab.Close;
             STListTab.ParamByName ('ID').ASInteger :=JPosTabARTIKEL_ID.Value;
             STListTab.Open;

             while not STListTab.Eof do
             begin
                // Menge erniedrigen
                ArtMengeTab.Close;
                ArtMengeTab.ParamByName ('ID').Value :=STListTabART_ID.Value;
                ArtMengeTab.ParamByName ('SUBMENGE').Value :=JPosTabMenge.Value * STListTabMENGE.Value;
                //Bestellmenge nicht verändern
                //ArtMengeTab.ParamByName ('BMENGE').Value :=0;
                ArtMengeTab.ExecSql;

                STListTab.Next;
             end;
             STListTab.Close;

             JPosTabGebucht.Value :=True;
          end; }

          // Daten aktualisieren
          //JPosTabJahr.Value :=Ja;
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := VK_RECH;
          JPosTabQuelle_Sub.Value := 2; // Kasse

          JPosTab.Post;
          JPosTab.Next;
        END;

      // im Artkikel die MENGE_VKRE_EDI aktualisieren
      UpdateArtikelEdiMenge(VK_RECH_EDI, JPosTabArtikel_ID.AsInteger, 0);

      //Seriennumern aktualisieren
      TRY
        //NEU
        UniQuery.Close;
        UniQuery.Sql.Text :=
          'select JPS.QUELLE,JPS.JOURNAL_ID,JPS.JOURNALPOS_ID,' +
          'JPS.ARTIKEL_ID,JPS.SNUM_ID from JOURNALPOS as JP, ' +
          'JOURNALPOS_SERNUM as JPS where JP.JOURNAL_ID=' +
          IntToStr(JourTabRec_ID.AsInteger) +
          ' and JP.SN_FLAG="Y" and JP.MENGE>0 and ' +
          'JP.ARTIKEL_ID=JPS.ARTIKEL_ID and ' +
          'JP.REC_ID=JPS.JOURNALPOS_ID and ' +
          'JP.JOURNAL_ID=JPS.JOURNAL_ID';

        SNSql := '';
        UniQuery.Open;
        WHILE NOT UniQuery.Eof DO
          BEGIN
            SNSql := SNSql +
              'UPDATE ARTIKEL_SERNUM SET STATUS="VK_RECH" ' +
              'WHERE ARTIKEL_ID=' +
              IntToStr(Uniquery.FieldByName('ARTIKEL_ID').AsInteger) +
              ' and SNUM_ID=' +
              IntToStr(Uniquery.FieldByName('SNUM_ID').AsInteger) + ';';
            DM1.UniQuery.Next;
          END;
        UniQuery.Close;

        ZBatchSql1.Sql.Text :=
          'UPDATE JOURNALPOS_SERNUM SET QUELLE=3 ' +
          'where JOURNAL_ID=' + IntToStr(JourTabRec_ID.AsInteger) + ';';

        IF length(SNSql) > 0 THEN
          ZBatchSql1.Sql.Add(SNSql);

        ProgressForm.Init(_('Seriennummern aktualisieren'));
        TRY
          ZBatchSql1.ExecSql;
        FINALLY
          ProgressForm.Stop;
        END;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der Seriennummern.'), mterror,
          [mbok], 0);
      END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen der Rechnung !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Angebot(Journal_ID : Integer) : STRING; // liefert AGB-Nummer zurück
VAR
  Pos : Integer;
BEGIN
  Result := '';
  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      // neue Nummer holen
      JourTabVRENUM.Value := IncNummerStr(VK_AGB);

      Result := JourTabVRENUM.Value;

      JourTabSTADIUM.Value := 0;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := VK_AGB;

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabFreigabe1_Flag.Value := False;

      // Kundendaten (Zahlungsart und Lieferart) aktualisieren,
      // falls diese noch nicht zugewiesen sind
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;
      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKun_Zahlart.AsInteger < 0 THEN
            KunTabKun_Zahlart.Value := JourTabZahlart.Value;
          IF KunTabKun_Liefart.AsInteger < 0 THEN
            KunTabKun_Liefart.Value := JourTabLiefart.Value;
          KunTab.Post;
        END;

      KunTab.Close;

      // Wenn MWST-Freies Angebot dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Daten aktualisieren
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := VK_AGB;

          JPosTab.Post;
          JPosTab.Next;
        END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen des Angebotes !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Auftrag(Journal_ID : Integer) : STRING; // liefert Auftrags-Nummer zurück
VAR
  Pos : Integer;
BEGIN
  Result := '';
  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      // neue Nummer holen
      JourTabVRENUM.Value := IncNummerStr(VK_AGB);

      Result := JourTabVRENUM.Value;

      JourTabSTADIUM.Value := 22;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := VK_Auftrag;

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabFreigabe1_Flag.Value := False;

      // Kundendaten (Zahlungsart und Lieferart) aktualisieren,
      // falls diese noch nicht zugewiesen sind
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;
      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKun_Zahlart.AsInteger < 0 THEN
            KunTabKun_Zahlart.Value := JourTabZahlart.Value;
          IF KunTabKun_Liefart.AsInteger < 0 THEN
            KunTabKun_Liefart.Value := JourTabLiefart.Value;
          KunTab.Post;
        END;

      KunTab.Close;

      // Wenn MWST-Freies Angebot dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Daten aktualisieren
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := VK_AGB;

          JPosTab.Post;
          JPosTab.Next;
        END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen des Angebotes !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;

//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_EKBest(Journal_ID : Integer) : STRING; // liefert EK-BST-Nummer zurück
VAR
  Pos : Integer;
  NewEK : Double;
BEGIN
  Result := '';
  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      // neue Nummer holen
      JourTabVRENUM.Value := IncNummerStr(EK_BEST);

      Result := JourTabVRENUM.Value;

      JourTabSTADIUM.Value := 0;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := EK_BEST;

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabFreigabe1_Flag.Value := False;

      JourTabStadium.AsInteger := 20;

      // Wenn MWST-Freie Bestellung dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      // Lieferantendaten (Zahlungsart) aktualisieren,
      // falls diese noch nicht zugewiesen ist
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;
      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKrd_Num.Value < 1 THEN
            BEGIN
              // Wenn noch keine KRD-NUM, dann neue Nummer zuweisen
              KunTabKrd_Num.Value := IncNummer(KRD_NUM_KEY);
              JourTabGEGENKONTO.Value := KunTabKrd_Num.Value;
              // Bitcodiertes Flag für "ist Lieferant" setzen
              KunTabSTATUS.AsInteger := KunTabSTATUS.AsInteger OR 16;
            END;
          IF KunTabLief_Zahlart.AsInteger < 0 THEN
            KunTabLief_Zahlart.Value := JourTabZahlart.Value;
          KunTab.Post;
        END;

      JourTab.Post;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Artikel Buchen  (Bestellmenge  erhöhen)
          IF (JPosTabArtikelTyp.Value = 'N') AND
            (JPosTabARTIKEL_ID.Value > -1) THEN
            BEGIN
              {
              // Menge nicht verändern !!!
              ArtMengeTab.Close;
              ArtMengeTab.ParamByName ('ID').Value :=JPosTabARTIKEL_ID.Value;
              ArtMengeTab.ParamByName ('SUBMENGE').Value :=0;
              //Bestellmenge erhöhen
              ArtMengeTab.ParamByName ('BMENGE').Value :=JPosTabMenge.Value*-1; // Menge wird abgezogen, daher - ( - * - = +)
              ArtMengeTab.ExecSql;
              JPosTabGebucht.Value :=True;
              }

              // Lieferanten im Artikelstamm aktualisieren

              NewEK := JPosTabEPREIS.AsFloat;

              IF JPosTabRabatt.Value <> 0 THEN
                BEGIN
                  IF JPosTabRabatt.Value = 100 THEN
                    NewEK := 0
                  ELSE
                    NewEK := NewEK - (NewEK * JPosTabRabatt.Value / 100);
                END;

              //Lieferantenpreis erstellen bzw. aktualisieren
              UpdateArtikelPreis(EK_RECH,
                JPosTabARTIKEL_ID.Value,
                JourTabAddr_ID.Value,
                NewEK);
            END;

          // Daten aktualisieren
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := EK_BEST;

          JPosTab.Post;

          // im Artkikel die MENGE_VKRE_EDI aktualisieren
          //UpdateArtikelEdiMenge (EK_BEST_EDI, JPosTabArtikel_ID.AsInteger,0);

          JPosTab.Next;
        END;

      JPosTab.Close;
      JourTab.Close;

      // EDI-Mengen in der Tabelle ARTIKEL_BDATEN aktualisieren
      UpdateArtikelEdiMenge(EK_BEST_EDI, JPosTabArtikel_ID.AsInteger, 0);
      // Menge bestellter Artikel aktualisieren
      UpdateEKBestMenge;

      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen der EK-Bestellung !'), mterror, [mbok],
        0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Einkauf(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
VAR
  NewEK,
    NewMenge,
    OldEK,
    OldMenge,
    //BM,ABM,
  Faktor,
    FaktorWGR,
    N, B : Double;

  UseRabGrp : Boolean;
  STA, OFF,
    SUM, I : Integer;
  LastID : Integer;
  Sql : STRING;
BEGIN
  Result := '';
  sum := 0;
  off := 0;
  //     UseRabGrp :=False;
  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      // neue Rechnungsnummer holen
      JourTabVRENUM.Value := IncNummerStr(EK_RECH);

      Result := JourTabVRENUM.Value;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabMAHNKOSTEN.Value := 0;
      JourTabFreigabe1_Flag.Value := False;

      // Lieferantendaten (Zahlungsart) aktualisieren,
      // falls diese noch nicht zugewiesen ist
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;
      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKrd_Num.Value < 1 THEN
            BEGIN
              // Wenn noch keine KRD-NUM, dann neue Nummer zuweisen
              KunTabKrd_Num.Value := IncNummer(KRD_NUM_KEY);
              JourTabGEGENKONTO.Value := KunTabKrd_Num.Value;
              // Bitcodiertes Flag für "ist Lieferant" setzen
              KunTabSTATUS.AsInteger := KunTabSTATUS.AsInteger OR 16;
            END;
          IF KunTabLief_Zahlart.AsInteger < 0 THEN
            KunTabLief_Zahlart.Value := JourTabZahlart.Value;
          KunTab.Post;
        END;

      CASE JourTabZahlart.Value OF
        //bar bzw. scheck
        1, 5 :
          BEGIN
            IF JourTabSOLL_SKONTO.Value > 0 THEN
              BEGIN
                JourTabStadium.Value := 80 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := JourTabSOLL_SKONTO.Value;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value :=
                  JourTabBSumme.Value - (JourTabBSumme.Value / 100) *
                  JourTabSOLL_SKONTO.Value;
              END
            ELSE
              BEGIN
                JourTabStadium.Value := 90 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := 0;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value := JourTabBSumme.Value;
              END;
          END;
        //Überweisung, NN, Lastschrift, EC-Karte
        2, 3, 4, 6, 9 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
      END;

      JourTabQuelle.Value := EK_Rech;

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;
      JourTabFREIGABE1_Flag.Value := False;

      // Wenn MWST-Freier Einkauf dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      IF JourTabZahlArt.Value = 1 THEN // Kassenbuchung
        BEGIN
          BucheKasse(JourTabIST_Zahldat.Value,
            EK_RECH,
            JourTabRec_ID.Value,
            JourTabVReNum.Value,
            JourTabGegenKonto.Value,
            JourTabIST_Skonto.Value,
            JourTabIST_Betrag.Value * -1, // negativ, da Ausgabe !!!
            'ZA EK-RE ' + JourTabKun_Name1.Value);
        END;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          TRY
            JPosTab.Edit;

            //MWST_Code löschen, wenn MwSt-freier Beleg
            IF JourTabMWST_FREI_FLAG.AsBoolean THEN
              JPosTabSteuer_Code.Value := 0;

            // EK-Preis, letzter EK, letzt. Lief, Lief-Datum
            // im Artikelstam aktualisieren !!!
            UpdateArtTab.Close;
            UpdateArtTab.ParamByName('ID').AsInteger := JPosTabARTIKEL_ID.Value;
            UpdateArtTab.Open;
            IF UpdateArtTab.RecordCount = 1 THEN
              BEGIN
                //UseRabGrp :=length(UpdateArtTab.FieldByName ('RABGRP_ID').AsString)>0;
                UseRabGrp :=
                  (length(UpdateArtTab.FieldByName('RABGRP_ID').AsString) > 0) AND
                  (UpdateArtTab.FieldByName('RABGRP_ID').AsString <> '-');

                TRY
                  UpdateArtTab.Edit;

                  OldEK := UpdateArtTab.FieldByName('EK_PREIS').AsFloat;
                  OldMenge := UpdateArtTab.FieldByName('MENGE_AKT').AsFloat;

                  NewEK := JPosTabEPREIS.AsFloat;
                  NewMenge := JPosTabMenge.AsFloat;

                  IF JPosTabRabatt.Value <> 0 THEN
                    BEGIN
                      IF JPosTabRabatt.Value = 100 THEN
                        NewEK := 0
                      ELSE
                        NewEK := NewEK - (NewEK * JPosTabRabatt.Value / 100);
                    END;

                  NewMenge := CAO_round(JPosTabMenge.AsFloat * 100) / 100; // auf 2 Nachkommastellen

                  //UpdateArtTab.FieldByName ('LAST_EK').AsFloat  :=CAO_round(NewEK*1000)/1000; // auf 3 Nachkommastellen
                  //UpdateArtTab.FieldByName ('LAST_LIEF').AsInteger :=JourTabAddr_ID.Value;
                  //UpdateArtTab.FieldByName ('LAST_LIEFDAT').AsDateTime :=JourTabRDatum.AsDateTime;

                  // Lieferantenpreis anlegen bzw. aktualisieren
                  UpdateArtikelPreis(EK_RECH,
                    JPosTabARTIKEL_ID.Value,
                    JourTabAddr_ID.Value,
                    NewEK);

                  // EK Berechnen (Mittelwert)
                  IF (OldMenge + NewMenge <> 0) AND (OldMenge >= 0) THEN
                    NewEK := (NewEK * NewMenge + OldEK * OldMenge) / (OldMenge +
                      NewMenge);

                  // nur Speichern, wenn keine Rabattgruppe gesetzt
                  IF NOT UseRabGrp THEN
                    BEGIN
                      UpdateArtTab.FieldByName('EK_PREIS').AsFloat :=
                        CAO_round(NewEK * 1000) / 1000; // auf 3 Nachkommastellen

                      // Berechnung von VK-Preisen wenn Kalkulationsfaktoren verwendet werden
                      FOR i := 1 TO AnzPreis DO
                        BEGIN
                          Faktor := GCalcFaktorTab[i];
                          FaktorWgr := 0;
                          GetWGRCalcFaktor(UpdateArtTab.FieldByName('WARENGRUPPE').AsInteger, i, FaktorWgr);
                          IF FaktorWgr <> 0 THEN
                            Faktor := FaktorWgr;

                          IF (Faktor > 0) THEN
                            BEGIN
                              N := CAO_round(NewEK * Faktor * 100) / 100;

                              UpdateArtTab.FieldByName('VK' +
                                IntToStr(i)).AsFloat := N;

                              B := N * (100 +
                                MwStTab[UpdateArtTab.FieldByName('STEUER_CODE').AsInteger]);

                              B := CAO_Round(B / BR_RUND_WERT) * BR_RUND_WERT /
                                100;

                              UpdateArtTab.FieldByName('VK' + IntToStr(i) +
                                'B').AsFloat := B
                            END;
                        END;
                    END;

                  //Bestelmenge aus Artikel holen
                  //ABM :=UpdateArtTab.FieldByName ('MENGE_BESTELLT').AsFloat;

                  UpdateArtTab.Post;
                EXCEPT
                  UpdateArtTab.Cancel;
                  MessageDlg(_('Fehler beim aktualisieren der Artikel-Mengen.'),
                    mterror, [mbok], 0);
                END;
              END;
            //else ABM :=999999;

            UpdateArtTab.Close;

            // Artikel Buchen
            IF (JPosTabGebucht.Value = False) AND
              (JPosTabArtikelTyp.Value = 'N') AND
              (JPosTabARTIKEL_ID.Value > -1) THEN
              BEGIN
                {
                if JPosTabQUELLE_SRC.AsInteger>0 then
                begin

                  // Bestellmenge aktualisieren
                  BM :=JPosTabMenge.AsFloat;

                  //Nachschauen, ob die Menge größer als die Best.-Menge ist und ggf. die Menge vermindern
                  uniquery.close;
                  uniquery.sql.text :='select MENGE from JOURNALPOS where QUELLE=6 and REC_ID='+IntToStr(JPosTabQUELLE_SRC.AsInteger);
                  uniquery.open;

                  if (UniQuery.RecordCount=1) then
                  begin
                     if BM > UniQuery.FieldByName('MENGE').AsFloat
                      then BM :=UniQuery.FieldByName('MENGE').AsFloat;
                  end else BM :=0;

                  uniquery.close;

                  // Artikel-Bestellmenge darf nicht < 0 sein !!!
                  if ABM-BM<0 then BM :=ABM;

                end else BM :=0; // war kein Artikel aus einer Bestellung
                }

                // Lager-Menge erhöhen
                ArtMengeTab.Close;
                ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
                ArtMengeTab.ParamByName('SUBMENGE').Value :=
                  (CAO_round(JPosTabMenge.Value * 100) / 100) * -1;
                //Bestellmenge erniedrigen
                //ArtMengeTab.ParamByName ('BMENGE').Value :=BM; // Menge wird abgezogen
                ArtMengeTab.ExecSql;
                JPosTabGebucht.Value := True;
              END;

            // Daten aktualisieren
            JPosTabVRENUM.Value := JourTabVRENUM.Value;
            JPosTabQuelle.Value := EK_RECH;

            JPosTab.Post;
          EXCEPT
            JPosTab.Cancel;
            MessageDlg(_('Fehler beim aktualisieren der EK-Positionen'), mterror,
              [mbok], 0);
          END;

          //Seriennumern aktualisieren
          IF JPosTabSN_Flag.AsBoolean THEN
            BEGIN
              TRY
                UniQuery.close;
                UniQuery.sql.text :=
                  'SELECT SNUM_ID from JOURNALPOS_SERNUM ' +
                  'WHERE ' +
                  'JOURNAL_ID=' + IntToStr(JourTabRec_ID.AsInteger) + ' and ' +
                  'JOURNALPOS_ID=' + IntToStr(JPosTabRec_ID.AsInteger);

                Sql := '';
                UniQuery.open;
                WHILE NOT UniQuery.Eof DO
                  BEGIN
                    SQL := SQL + 'UPDATE ARTIKEL_SERNUM SET STATUS="LAGER" ' +
                      'WHERE ' +
                      'ARTIKEL_ID=' + IntToStr(JPosTabArtikel_ID.AsInteger) +
                      ' and ' +
                      'SNUM_ID=' +
                      IntToStr(UniQuery.FieldByName('SNUM_ID').AsInteger) +
                      ' and STATUS="EK_EDI";';

                    UniQuery.Next;
                  END;
                UniQuery.Close;

                IF length(SQL) > 0 THEN
                  BEGIN
                    ZBatchSql1.Sql.Text := Sql;
                    ProgressForm.Init(_('Seriennummern aktualisieren'));
                    TRY
                      ZBatchSql1.ExecSql;
                    FINALLY
                      ProgressForm.Stop;
                    END;
                  END;
              EXCEPT
                MessageDlg(_('Fehler beim aktualisieren der Seriennummern. (1)'), mterror, [mbok], 0);
              END;
            END;

          JPosTab.Next;
        END;

      //Seriennumern aktualisieren
      TRY
        UniQuery.close;
        UniQuery.sql.text :=
          'UPDATE JOURNALPOS_SERNUM SET QUELLE=5 WHERE ' +
          'JOURNAL_ID=' + IntToStr(JourTabRec_ID.AsInteger) +
          ' and QUELLE=15';

        UniQuery.ExecSql;
        UniQuery.close;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der Seriennummern. (2)'),
          mterror, [mbok], 0);
      END;

      // im Artkikel die MENGE_VKRE_EDI aktualisieren
      UpdateArtikelEdiMenge(EK_RECH_EDI, JPosTabArtikel_ID.AsInteger, 0);
      // EK-Bestellmenge aktualisieren
      UpdateEKBestMenge;

      // Status der Bestellungen dieses Lieferanten aktualisieren
      uniquery.close;
      uniquery.sql.clear;
      uniquery.sql.add('select');
      uniquery.sql.add('JOURNAL.REC_ID,SUM(JP2.MENGE) as ');
      uniquery.sql.add('MENGE_EK,JOURNALPOS.MENGE as MENGE_BEST');
      uniquery.sql.add('from (JOURNALPOS ,JOURNAL)');
      uniquery.sql.add('left outer join JOURNALPOS as JP2 on ');
      uniquery.sql.add('JP2.QUELLE_SRC = JOURNALPOS.REC_ID and JP2.QUELLE<>15');
      uniquery.sql.add('where JOURNAL.QUELLE=6 and JOURNALPOS.QUELLE=6 and ');
      uniquery.sql.add('JOURNALPOS.ADDR_ID=' + Inttostr(JourTabAddr_ID.AsInteger));
      uniquery.sql.add('and (JOURNALPOS.ARTIKELTYP="N" or ');
      uniquery.sql.add('JOURNALPOS.ARTIKELTYP="S") and ');
      uniquery.sql.add('JOURNAL.REC_ID=JOURNALPOS.JOURNAL_ID and ');
      uniquery.sql.add('JOURNAL.STADIUM>=20 and JOURNAL.STADIUM<100');
      uniquery.sql.add('group by JOURNALPOS.REC_ID');

      LastID := -1;
      TRY
        uniquery.open;

        WHILE NOT uniquery.eof DO
          BEGIN
            IF LastID <> UniQuery.FieldByName('REC_ID').AsInteger THEN
              BEGIN
                IF LastID <> -1 THEN
                  BEGIN
                    IF off = 0 THEN
                      STA := 100
                    ELSE
                      IF off = sum THEN
                        STA := 20
                      ELSE
                        STA := 30;

                    uniquery2.close;
                    uniquery2.sql.text := 'UPDATE JOURNAL SET STADIUM=' +
                      Inttostr(STA) +
                      ' WHERE QUELLE=6 and REC_ID=' +
                      Inttostr(LastID);
                    uniquery2.execsql;
                  END;
                sum := 0;
                off := 0;
                LastID := UniQuery.FieldByName('REC_ID').AsInteger;
              END;
            sum := sum + UniQuery.FieldByName('MENGE_BEST').AsInteger;
            off := off + (UniQuery.FieldByName('MENGE_BEST').AsInteger -
              UniQuery.FieldByName('MENGE_EK').AsInteger);

            uniquery.next;
          END;
        IF LastID <> -1 THEN
          BEGIN
            IF off < 1 THEN
              STA := 100
            ELSE
              IF off = sum THEN
                STA := 20
              ELSE
                STA := 30;

            uniquery2.close;
            uniquery2.sql.text := 'UPDATE JOURNAL SET STADIUM=' +
              Inttostr(STA) +
              ' WHERE QUELLE=6 and REC_ID=' +
              Inttostr(LastID);
            uniquery2.execsql;
          END;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der EK-Bestellungen.'), mterror,
          [mbok], 0);
      END;
      uniquery.close;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen der EK-Rechnung !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Gutschrift(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
//var  NewEK,
//    NewMenge,
//    OldEK,
//    OldMenge,
    //BM,ABM,
//    Faktor,
//    FaktorWGR : Double;

//    UseRabGrp : Boolean;
//    STA,OFF,
//    SUM, I    : Integer;
//    LastID    : Integer;
//    Sql       : String;
BEGIN
  Result := '';
  //     sum    := 0;
  //     off    := 0;
  //     UseRabGrp :=False;
  Transact1.AutoCommit := False;
  Transact1.TransactSafe := True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      // neue Rechnungsnummer holen
      JourTabVRENUM.Value := IncNummerStr(Gutschrift);
      Result := JourTabVRENUM.Value;

      JourTabStadium.Value := 50;
      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabMAHNKOSTEN.Value := 0;
      JourTabQuelle.Value := Gutschrift;
      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;
      JourTabFREIGABE1_Flag.Value := False;
      JourTabMWST_0.AsInteger := 0;
      JourTabMWST_1.AsInteger := 0;
      JourTabMWST_2.AsInteger := 0;
      JourTabMWST_3.AsInteger := 0;
      JourTabMSUMME_0.AsFloat := 0;
      JourTabMSUMME_1.AsFloat := 0;
      JourTabMSUMME_2.AsFloat := 0;
      JourTabMSUMME_3.AsFloat := 0;
      JourTabMSUMME.AsFloat := 0;
      JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;

      JourTab.Post;
      // JourTab.Close;

      Transact1.Commit;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          TRY
            JPosTab.Edit;

            JPosTabSteuer_Code.Value := 0;
            // neuen Artikel anfügen
            (*
            IF (JPosTabGebucht.Value = False) AND
              (JPosTabARTIKEL_ID.Value < -1) THEN
              BEGIN
                TRY
                  UniQuery2.Close;
                  UniQuery2.Sql.Clear;
                  UniQuery2.Sql.Text := 'insert into artikel set ' +
                    'Artnum = "' + JPosTabArtNum.Value + '",' +

                  'Kurzname = "' + stringtoSql(JPosTabBezeichnung.Value) + '",'
                    +
                    'EK_Preis = "' + JPosTabEPreis.AsString + '",' +
                    'Einlieferer_Name = "' + JourTabKun_Name1.Value + '",' +
                    'Einlieferer_ID = "' + JourTabADDR_ID.AsString + '",' +
                    'Einlieferer_Nr = "' + JourTabKun_Num.AsString + '"'

                  ;

                                      'Artnum = "'+ JPosTabArtNum.Value+'",'+
                    (artnum,kurzname,EK_preis,einlieferer_name,'+
                                     'einlieferer_ID,einlieferer_Nummer) values ('+
                                      '"'+JPosTabArtNum.Value+'",'+
                                      '"'+JPosTabBEZEICHNUNG.Value+'",'+
                                      '"'+JposTabEPreis.AsString + '",'+
                                      '"'+JourTabKun_Name1.Value+'",'+
                                      '"'+JourTabADDR_ID.AsString +'",'+
                                      '"'+JourTabKun_Num.Value+'")');

                  UniQuery2.ExecSql;
                  JPosTabGebucht.Value := True;
                  // Daten aktualisieren
                  JPosTabVRENUM.Value := JourTabVRENUM.Value;
                  JPosTabQuelle.Value := Gutschrift;
                  UniQuery2.Sql.Clear;
                  UniQuery2.Sql.Text :=
                    'select Rec_ID from Artikel Where ArtNum = "' +
                    JPosTabArtNum.Value + '"';
                  UniQuery2.Open;
                  JPosTabARTIKEL_ID.Value :=
                    UniQuery2.FieldByName('REC_ID').AsInteger;
                  JPosTab.Post;
                EXCEPT
                  MessageDlg(_('Fehler beim anfügen des Artikels'), mterror,
                    [mbok], 0);
                END;
              END; *)

            //JPosTab.Cancel;
            //MessageDlg (_('Fehler beim aktualisieren der EK-Positionen'),mterror,[mbok],0);

            JPosTab.Next;
          EXCEPT
            MessageDlg(_('Fehler beim Aktualisieren der Position'), mterror,
              [mbok], 0);
          END;
        END; // while not eof
    EXCEPT
      JPosTab.Close;
      JourTab.Close;
      UniQuery2.Close;
      Transact1.Commit;

      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    Transact1.TransactSafe := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Abrechnung(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
VAR
  Pos : Integer;
  //IStr     : String;
  SNSql : STRING;
BEGIN
  Result := '';

  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      JourTabVRENUM.Value := IncNummerStr(Abrechnung);
      Result := JourTabVRENUM.Value;
      JourTabSTADIUM.Value := 50;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := Abrechnung;
      JourTabQuelle_Sub.Value := 1; // Rechnung

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabFreigabe1_Flag.Value := False;
      JourTabMAHNKOSTEN.Value := 0;

      // Kundendaten (Zahlungsart und Lieferart) aktualisieren,
      // falls diese noch nicht zugewiesen sind
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;

      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKun_Zahlart.AsInteger < 0 THEN
            KunTabKun_Zahlart.Value := JourTabZahlart.Value;
          IF KunTabKun_Liefart.AsInteger < 0 THEN
            KunTabKun_Liefart.Value := JourTabLiefart.Value;
          KunTab.Post;
        END;

      KunTab.Close;

      CASE JourTabZahlart.Value OF
        //bar bzw. scheck
        1, 5 :
          BEGIN
            IF JourTabSOLL_SKONTO.Value > 0 THEN
              BEGIN
                JourTabStadium.Value := 80 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := JourTabSOLL_SKONTO.Value;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value :=
                  JourTabBSumme.Value - (JourTabBSumme.Value / 100) *
                  JourTabSOLL_SKONTO.Value;
              END
            ELSE
              BEGIN
                JourTabStadium.Value := 90 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := 0;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value := JourTabBSumme.Value;
              END;
          END;
        //Überweisung
        2, 3, 4 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
        //Lastschrift, EC-Karte
        6, 9 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
      END;

      // Wenn MWST-Freie Rechnung dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      IF JourTabZahlArt.Value = 1 THEN // Kassenbuchung
        BEGIN
          BucheKasse(JourTabIST_Zahldat.Value,
            VK_RECH, JourTabRec_ID.Value,
            JourTabVReNum.Value,
            JourTabGegenKonto.Value,
            JourTabIST_Skonto.Value,
            JourTabIST_Betrag.Value,
            'ZE VK-RE ' + JourTabKun_Name1.Value);

        END;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      //IStr :=''; // Insert-String für Stücklistenartikel leeren

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Artikel Buchen
          IF (JPosTabGebucht.Value = False) AND
            (
            (JPosTabArtikelTyp.Value = 'N') OR
            (JPosTabArtikelTyp.Value = 'X')
            ) AND
            (JPosTabARTIKEL_ID.Value > -1) THEN
            BEGIN
              // Menge erniedrigen
              ArtMengeTab.Close;
              ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
              ArtMengeTab.ParamByName('SUBMENGE').Value := JPosTabMenge.Value;
              ArtMengeTab.ExecSql;

              JPosTabGebucht.Value := True;
            END;

          // Daten aktualisieren
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := Abrechnung;
          JPosTabQuelle_Sub.Value := 1; // Rechnung

          JPosTab.Post;

          // im Artkikel als abgerechnet markieren
          Uniquery.Close;
          UniQuery.SQL.Text :=
            'Update Artikel set abgerechnet = 1  where Rec_id = ' +
            IntToStr(JPosTabARTIKEL_ID.AsInteger);
          Uniquery.ExecSql;
          Uniquery.close;

          JPosTab.Next;
        END;

      //Seriennumern aktualisieren
      TRY
        {
        dm1.UniQuery.close;
        dm1.UniQuery.sql.text :='UPDATE ARTIKEL_SERNUM SET VERK_NUM='+
                                IntToStr(JourTabVRENUM.AsInteger)+
                                ' where VK_JOURNAL_ID='+
                                IntToStr(JourTabRec_ID.AsInteger);
        dm1.UniQuery.ExecSql;
        dm1.UniQuery.close;  }

        //NEU
        UniQuery.Close;
        UniQuery.Sql.Text :=
          'select JPS.QUELLE,JPS.JOURNAL_ID,JPS.JOURNALPOS_ID,' +
          'JPS.ARTIKEL_ID,JPS.SNUM_ID from JOURNALPOS as JP, ' +
          'JOURNALPOS_SERNUM as JPS where JP.JOURNAL_ID=' +
          IntToStr(JourTabRec_ID.AsInteger) +
          ' and JP.SN_FLAG="Y" and JP.MENGE>0 and ' +
          'JP.ARTIKEL_ID=JPS.ARTIKEL_ID and ' +
          'JP.REC_ID=JPS.JOURNALPOS_ID and ' +
          'JP.JOURNAL_ID=JPS.JOURNAL_ID';

        SNSql := '';
        UniQuery.Open;
        WHILE NOT UniQuery.Eof DO
          BEGIN
            SNSql := SNSql +
              'UPDATE ARTIKEL_SERNUM SET STATUS="VK_RECH" ' +
              'WHERE ARTIKEL_ID=' +
              IntToStr(Uniquery.FieldByName('ARTIKEL_ID').AsInteger) +
              ' and SNUM_ID=' +
              IntToStr(Uniquery.FieldByName('SNUM_ID').AsInteger) + ';';
            UniQuery.Next;
          END;
        UniQuery.Close;

        ZBatchSql1.Sql.Text :=
          'UPDATE JOURNALPOS_SERNUM SET QUELLE=3 ' +
          'where JOURNAL_ID=' + IntToStr(JourTabRec_ID.AsInteger) + ';';

        IF length(SNSql) > 0 THEN
          ZBatchSql1.Sql.Add(SNSql);

        ProgressForm.Init(_('Seriennummern aktualisieren'));
        TRY
          ZBatchSql1.ExecSql;
        FINALLY
          ProgressForm.Stop;
        END;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der Seriennummern.'),
          mterror, [mbok], 0);
      END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen der Rechnung !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;

//------------------------------------------------------------------------------

FUNCTION tDM1.Buche_Rueckgabe(Journal_ID : Integer) : STRING; // liefert Belegnummer zurück
VAR
  Pos : Integer;
  //IStr     : String;
  SNSql : STRING;
BEGIN
  Result := '';

  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := Journal_ID;
      JourTab.Open;
      JourTab.Edit;
      JourTabVRENUM.Value := IncNummerStr(Rueckgabe);
      Result := JourTabVRENUM.Value;
      JourTabSTADIUM.Value := 50;
      JourTabRDatum.Value := now;
      JourTabQuelle.Value := Rueckgabe;
      JourTabQuelle_Sub.Value := 1; // Rechnung

      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      JourTabIST_Betrag.Value := 0;
      JourTabIST_Anzahlung.Value := 0;
      JourTabFreigabe1_Flag.Value := False;
      JourTabMAHNKOSTEN.Value := 0;

      // Kundendaten (Zahlungsart und Lieferart) aktualisieren,
      // falls diese noch nicht zugewiesen sind
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;

      IF KunTab.RecordCount = 1 THEN
        BEGIN
          KunTab.Edit;
          IF KunTabKun_Zahlart.AsInteger < 0 THEN
            KunTabKun_Zahlart.Value := JourTabZahlart.Value;
          IF KunTabKun_Liefart.AsInteger < 0 THEN
            KunTabKun_Liefart.Value := JourTabLiefart.Value;
          KunTab.Post;
        END;

      KunTab.Close;

      CASE JourTabZahlart.Value OF
        //bar bzw. scheck
        1, 5 :
          BEGIN
            IF JourTabSOLL_SKONTO.Value > 0 THEN
              BEGIN
                JourTabStadium.Value := 80 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := JourTabSOLL_SKONTO.Value;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value :=
                  JourTabBSumme.Value - (JourTabBSumme.Value / 100) *
                  JourTabSOLL_SKONTO.Value;
              END
            ELSE
              BEGIN
                JourTabStadium.Value := 90 + JourTabZahlart.Value;
                JourTabIST_SKONTO.Value := 0;
                JourTabIST_ANZAHLUNG.Value := 0;
                JourTabIST_ZAHLDAT.Value := JourTabRDatum.Value;
                JourTabIST_BETRAG.Value := JourTabBSumme.Value;
              END;
          END;
        //Überweisung
        2, 3, 4 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
        //Lastschrift, EC-Karte
        6, 9 : JourTabStadium.Value := 20 + JourTabZahlart.Value;
      END;

      // Wenn MWST-Freie Rechnung dann alle MWST-Felder löschen
      IF JourTabMWST_FREI_FLAG.AsBoolean THEN
        BEGIN
          JourTabMWST_0.AsInteger := 0;
          JourTabMWST_1.AsInteger := 0;
          JourTabMWST_2.AsInteger := 0;
          JourTabMWST_3.AsInteger := 0;
          JourTabMSUMME_0.AsFloat := 0;
          JourTabMSUMME_1.AsFloat := 0;
          JourTabMSUMME_2.AsFloat := 0;
          JourTabMSUMME_3.AsFloat := 0;
          JourTabMSUMME.AsFloat := 0;
          JourTabBSUMME.AsFloat := JourTabNSUMME.AsFloat;
        END;

      JourTab.Post;

      IF JourTabZahlArt.Value = 1 THEN // Kassenbuchung
        BEGIN
          BucheKasse(JourTabIST_Zahldat.Value,
            VK_RECH, JourTabRec_ID.Value,
            JourTabVReNum.Value,
            JourTabGegenKonto.Value,
            JourTabIST_Skonto.Value,
            JourTabIST_Betrag.Value,
            'ZE VK-RE ' + JourTabKun_Name1.Value);

        END;

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := Journal_ID;
      JPosTab.Open;

      Pos := 0;

      //IStr :=''; // Insert-String für Stücklistenartikel leeren

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          //MWST_Code löschen, wenn MwSt-freier Beleg
          IF JourTabMWST_FREI_FLAG.AsBoolean THEN
            JPosTabSteuer_Code.Value := 0;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF (JPosTabArtikelTyp.Value <> 'T') AND
            (JPosTabArtikelTyp.Value <> 'X') THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Artikel Buchen
          IF (JPosTabGebucht.Value = False) AND
            (
            (JPosTabArtikelTyp.Value = 'N') OR
            (JPosTabArtikelTyp.Value = 'X')
            ) AND
            (JPosTabARTIKEL_ID.Value > -1) THEN
            BEGIN
              // Menge erniedrigen
              ArtMengeTab.Close;
              ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
              ArtMengeTab.ParamByName('SUBMENGE').Value := JPosTabMenge.Value;
              ArtMengeTab.ExecSql;

              JPosTabGebucht.Value := True;
            END;

          // Daten aktualisieren
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := Abrechnung;
          JPosTabQuelle_Sub.Value := 1; // Rechnung

          JPosTab.Post;

          // im Artkikel als abgerechnet markieren
          Uniquery.Close;
          UniQuery.SQL.Text :=
            'Update Artikel set abgerechnet = 1  where Rec_id = ' +
            IntToStr(JPosTabARTIKEL_ID.AsInteger);
          Uniquery.ExecSql;
          Uniquery.close;

          JPosTab.Next;
        END;

      //Seriennumern aktualisieren
      TRY
        {
        dm1.UniQuery.close;
        dm1.UniQuery.sql.text :='UPDATE ARTIKEL_SERNUM SET VERK_NUM='+
                                IntToStr(JourTabVRENUM.AsInteger)+
                                ' where VK_JOURNAL_ID='+
                                IntToStr(JourTabRec_ID.AsInteger);
        dm1.UniQuery.ExecSql;
        dm1.UniQuery.close;  }

        //NEU
        UniQuery.Close;
        UniQuery.Sql.Text :=
          'select JPS.QUELLE,JPS.JOURNAL_ID,JPS.JOURNALPOS_ID,' +
          'JPS.ARTIKEL_ID,JPS.SNUM_ID from JOURNALPOS as JP, ' +
          'JOURNALPOS_SERNUM as JPS where JP.JOURNAL_ID=' +
          IntToStr(JourTabRec_ID.AsInteger) +
          ' and JP.SN_FLAG="Y" and JP.MENGE>0 and ' +
          'JP.ARTIKEL_ID=JPS.ARTIKEL_ID and ' +
          'JP.REC_ID=JPS.JOURNALPOS_ID and ' +
          'JP.JOURNAL_ID=JPS.JOURNAL_ID';

        SNSql := '';
        UniQuery.Open;
        WHILE NOT UniQuery.Eof DO
          BEGIN
            SNSql := SNSql +
              'UPDATE ARTIKEL_SERNUM SET STATUS="VK_RECH" ' +
              'WHERE ARTIKEL_ID=' +
              IntToStr(Uniquery.FieldByName('ARTIKEL_ID').AsInteger) +
              ' and SNUM_ID=' +
              IntToStr(Uniquery.FieldByName('SNUM_ID').AsInteger) + ';';
            UniQuery.Next;
          END;
        UniQuery.Close;

        ZBatchSql1.Sql.Text :=
          'UPDATE JOURNALPOS_SERNUM SET QUELLE=3 ' +
          'where JOURNAL_ID=' + IntToStr(JourTabRec_ID.AsInteger) + ';';

        IF length(SNSql) > 0 THEN
          ZBatchSql1.Sql.Add(SNSql);

        ProgressForm.Init(_('Seriennummern aktualisieren'));
        TRY
          ZBatchSql1.ExecSql;
        FINALLY
          ProgressForm.Stop;
        END;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der Seriennummern.'),
          mterror, [mbok], 0);
      END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen der Rechnung !'), mterror, [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;

//------------------------------------------------------------------------------
// Nur zur Probe aus der EDI-Rechnung einen Lieferschein erstellen !

FUNCTION tDM1.Buche_Lieferschein(Journal_ID : Integer;
  Teillief : Boolean;
  VAR LieferscheinID : Integer) : STRING; // liefert Belegnummer zurück
VAR
  NewId : Integer;
  Pos : Integer;
  //S        : String;
BEGIN
  IF NOT TeilLief THEN
    NewID := CopyRechnung(Journal_ID, VK_LIEF_EDI)
  ELSE
    NewID := Journal_ID;

  LieferscheinID := NewID;

  Result := '';
  Transact1.AutoCommit := False;
  //Transact1.TransactSafe :=True;
  TRY
    TRY
      JourTab.Close;
      JourTab.ParamByName('ID').Value := NewID;
      JourTab.Open;
      JourTab.Edit;

      // Nachschauen, ob Kunde eine Lieferanschrift hat
      KunTab.Close;
      KunTab.ParamByName('ID').AsInteger := JourTabAddr_ID.Value;
      KunTab.Open;

      // neu 22.08.2003
      // Lieferanschrift zuweisen
      WITH UniQuery DO
        BEGIN
          SQL.Text := 'select * from ADRESSEN_LIEF where REC_ID=' +
            IntToStr(JourTabLIEF_ADDR_ID.AsInteger);
          Open;
          IF RecordCount > 0 THEN
            BEGIN
              JourTabKUN_Anrede.Text := FieldByName('ANREDE').AsString;
              JourTabKUN_NAME1.Text := FieldByName('NAME1').AsString;
              JourTabKUN_NAME2.Text := FieldByName('NAME2').AsString;
              JourTabKUN_NAME3.Text := FieldByName('NAME3').AsString;
              JourTabKUN_Strasse.Text := FieldByName('STRASSE').AsString;
              JourTabKUN_LAND.Text := FieldByName('LAND').AsString;
              JourTabKUN_PLZ.Text := FieldByName('PLZ').AsString;
              JourTabKUN_Ort.Text := FieldByName('ORT').AsString;
            END
        END;

      // neue Belegnummer holen
      JourTabVLSNUM.AsString := IncNummerStr(VK_LIEF);
      JourTabVRENUM.AsString := '';
      JourTabRDatum.Value := 0;
      JourTabFreigabe1_Flag.Value := False;

      Result := JourTabVLSNUM.AsString;

      JourTabLDatum.Value := Now;
      JourTabStadium.Value := 20 + JourTabLiefart.Value;

      JourTabQuelle.Value := VK_LIEF;
      JourTabKONTOAUSZUG.Value := -1;
      JourTabUW_NUM.Value := -1;
      JourTabBANK_ID.Value := -1;

      // Warenwert erst mal auf Nettosumme !!!
      // Muß noch geändert werden !!!
      JourTabWERT_NETTO.Value := JourTabNSumme.Value;

      JourTab.Post;

      // Lieferscheinnummer in die Seriennummern eintragen
      {
      UniQuery.Sql.Text :=
         'UPDATE ARTIKEL_SERNUM SET LIEF_NUM='+
         Inttostr(JourTabVLSNUM.AsInteger)+
         ' WHERE LS_JOURNAL_ID='+
         IntToStr(NewID);
      UniQuery.ExecSql;
      }

      JPosTab.Close;
      JPosTab.ParamByName('ID').Value := NewID;
      JPosTab.Open;

      Pos := 0;

      WHILE NOT JPosTab.Eof DO
        BEGIN
          JPosTab.Edit;

          // Position schreiben, aber nur bei Artikeln, nicht bei Text
          IF JPosTabArtikelTyp.Value <> 'T' THEN
            BEGIN
              Inc(Pos);
              JposTabView_Pos.Value := Inttostr(Pos);
            END;

          // Daten aktualisieren
          JPosTabVLSNUM.Value := JourTabVLSNUM.Value;
          JPosTabVRENUM.Value := JourTabVRENUM.Value;
          JPosTabQuelle.Value := VK_LIEF;

          JPosTab.Post;
          JPosTab.Next;
        END;

      JPosTab.Close;
      JourTab.Close;
      Transact1.Commit;

    EXCEPT
      MessageDlg(_('Fehler beim Buchen des VK-Lieferscheines !'), mterror,
        [mbok], 0);
      Transact1.RollBack;
      Result := '';
    END;
  FINALLY
    Transact1.AutoCommit := True;
    //Transact1.TransactSafe :=False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Storno_Einkauf(Journal_ID : Integer) : Boolean; // True, Wenn OK
VAR
  S : STRING;
BEGIN
  //     Result :=False;
  TRY
    JourTab.Close;
    JourTab.ParamByName('ID').Value := Journal_ID;
    JourTab.Open;

    IF JourTabZahlArt.Value = 1 THEN // Kassenbuchung stornieren
      BEGIN
        WITH UniQuery DO
          BEGIN
            Close;
            SQL.Clear;
            SQL.Add('delete from FIBU_KASSE');
            SQL.Add('where JOURNAL_ID=' + Inttostr(Journal_ID));
            ExecSql;
            SQL.Clear;
          END;
      END;

    // ggf. Seriennummer freigeben
    WITH UniQuery DO
      BEGIN
        Close;
        SQL.Text := 'select JPS.*, ASN.SERNUMMER, ASN.STATUS ' +
          'from JOURNALPOS_SERNUM JPS, ARTIKEL_SERNUM ASN ' +
          'where JOURNAL_ID=' + Inttostr(Journal_ID) +
          ' and JPS.ARTIKEL_ID=ASN.ARTIKEL_ID and ' +
          'JPS.SNUM_ID=ASN.SNUM_ID and ASN.STATUS="LAGER"';
        Open;
        S := '';
        IF RecordCount > 0 THEN
          BEGIN
            WHILE NOT Eof DO
              BEGIN
                IF length(s) > 0 THEN
                  s := s + ',';
                s := s + IntToStr(FieldByName('SNUM_ID').AsInteger);
                Next;
              END;
          END;
        Close;

        IF length(s) > 0 THEN
          BEGIN
            SQL.Text := 'DELETE FROM ARTIKEL_SERNUM ' +
              'WHERE SNUM_ID IN (' + s + ')';
            ExecSql;
          END;

        Sql.Text := 'DELETE FROM JOURNALPOS_SERNUM where ' +
          'JOURNAL_ID=' + IntToStr(Journal_ID);
        ExecSql;

        {SQL.Clear;
        SQL.Add ('update ARTIKEL_SERNUM set EK_JOURNAL_ID=-1, ');
        SQL.Add ('EK_JOURNALPOS_ID=-1 where EK_JOURNAL_ID='+Inttostr(Journal_ID));
        ExecSql;     }
        SQL.Clear;
      END;

    JPosTab.Close;
    JPosTab.ParamByName('ID').Value := Journal_ID;
    JPosTab.Open;

    WHILE NOT JPosTab.Eof DO
      BEGIN
        // Artikel Buchen
        IF (JPosTabArtikelTyp.Value = 'N') AND
          (JPosTabARTIKEL_ID.Value > -1) THEN
          BEGIN
            // Menge erniedrugen (EK STORNO)
            ArtMengeTab.Close;
            ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
            ArtMengeTab.ParamByName('SUBMENGE').Value := JPosTabMenge.Value;
            //ArtMengeTab.ParamByName ('BMENGE').Value :=0;
            ArtMengeTab.ExecSql;
            JPosTabGebucht.Value := True;
          END;

        // Daten aktualisieren
        JPosTab.Delete;
      END;
    JPosTab.Close;

    JourTab.Delete;
    JourTab.Close;

    //Datei-Links löschen
    LinkForm.DelLinks(EK_RECH, JOURNAL_ID);

    Result := True;
  EXCEPT
    MessageDlg(_('Fehler beim Storno der EK-Rechnung !'), mterror, [mbok], 0);
    Result := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Storno_Verkauf(Journal_ID : Integer) : Boolean; // True, Wenn OK
VAR
  S : STRING;
BEGIN
  //     Result :=False;
  TRY
    JourTab.Close;
    JourTab.ParamByName('ID').Value := Journal_ID;
    JourTab.Open;

    IF JourTabZahlArt.Value = 1 THEN // Kassenbuchung stornieren
      BEGIN
        WITH UniQuery DO
          BEGIN
            Close;
            SQL.Clear;
            SQL.Add('DELETE FROM FIBU_KASSE');
            SQL.Add('WHERE JOURNAL_ID=' + Inttostr(Journal_ID));
            ExecSql;
            SQL.Clear;
          END;
      END;

    // ggf. Seriennummer freigeben
    WITH UniQuery DO
      BEGIN
        Close;
        //SQL.Clear;
        SQL.Text := 'SELECT * FROM JOURNALPOS_SERNUM ' +
          'WHERE JOURNAL_ID=' + InttoStr(Journal_ID);
        Open;
        S := '';
        IF RecordCount > 0 THEN
          BEGIN
            WHILE NOT Eof DO
              BEGIN
                IF length(s) > 0 THEN
                  s := s + ',';
                s := s + IntToStr(FieldByName('SNUM_ID').AsInteger);
                Next;
              END;
          END;
        Close;

        IF length(s) > 0 THEN
          BEGIN
            SQL.Text := 'UPDATE ARTIKEL_SERNUM SET STATUS="LAGER" ' +
              'WHERE SNUM_ID IN (' + s + ')';
            ExecSql;
          END;

        Close;
        SQL.Text := 'DELETE FROM JOURNALPOS_SERNUM ' +
          'WHERE JOURNAL_ID=' + InttoStr(Journal_ID);
        ExecSql;

        //SQL.Add ('update ARTIKEL_SERNUM set VK_JOURNAL_ID=-1, ');
        //SQL.Add ('VK_JOURNALPOS_ID=-1 where VK_JOURNAL_ID='+Inttostr(Journal_ID));
        //ExecSql;
        //SQL.Clear;
      END;

    JPosTab.Close;
    JPosTab.ParamByName('ID').Value := Journal_ID;
    JPosTab.Open;

    WHILE NOT JPosTab.Eof DO
      BEGIN
        // Artikel Buchen
        IF (JPosTabArtikelTyp.Value = 'N') AND
          (JPosTabARTIKEL_ID.Value > -1) THEN
          BEGIN
            // Menge erhöhen (VK STORNO)
            ArtMengeTab.Close;
            ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
            ArtMengeTab.ParamByName('SUBMENGE').Value := JPosTabMenge.Value *
              -1;
            //ArtMengeTab.ParamByName ('BMENGE').Value :=0;
            ArtMengeTab.ExecSql;
            //JPosTabGebucht.Value :=True;
          END
        ELSE
          IF (JPosTabArtikelTyp.Value = 'S') AND
            (JPosTabARTIKEL_ID.Value > -1) THEN
            BEGIN
              // Stückliste, Unterartikel Menge korregieren
              STListTab.Close;
              STListTab.ParamByName('ID').AsInteger := JPosTabARTIKEL_ID.Value;
              STListTab.Open;
              WHILE NOT STListTab.Eof DO
                BEGIN
                  // Menge erhöhen (VK STORNO)
                  ArtMengeTab.Close;
                  ArtMengeTab.ParamByName('ID').Value :=
                    STListTabART_ID.AsInteger;
                  ArtMengeTab.ParamByName('SUBMENGE').Value :=
                    STListTabMENGE.Value * JPosTabMenge.Value * -1;
                  //ArtMengeTab.ParamByName ('BMENGE').Value :=0;
                  ArtMengeTab.ExecSql;

                  STListTab.Next;
                END;
              STListTab.Close;
            END;

        // Verweise löschen
        UniQuery.Sql.Text := 'UPDATE JOURNALPOS SET QUELLE_SRC=-1 ' +
          'WHERE QUELLE_SRC=' +
          Inttostr(JPosTabRec_ID.AsInteger);
        UniQuery.ExecSql;

        // Daten aktualisieren
        //JPosTab.Delete;
        JPosTab.Edit;
        JPosTabGebucht.AsBoolean := False;
        JPosTab.Post;

        JPosTab.Next;
      END;
    JPosTab.Close;

    JourTab.Edit;

    JourTabStadium.AsInteger := 127;

    JourTabNSUMME_0.AsFloat := 0;
    JourTabNSUMME_1.AsFloat := 0;
    JourTabNSUMME_2.AsFloat := 0;
    JourTabNSUMME_3.AsFloat := 0;
    JourTabNSUMME.AsFloat := 0;

    JourTabMSUMME_0.AsFloat := 0;
    JourTabMSUMME_1.AsFloat := 0;
    JourTabMSUMME_2.AsFloat := 0;
    JourTabMSUMME_3.AsFloat := 0;
    JourTabMSUMME.AsFloat := 0;

    JourTabBSUMME_0.AsFloat := 0;
    JourTabBSUMME_1.AsFloat := 0;
    JourTabBSUMME_2.AsFloat := 0;
    JourTabBSUMME_3.AsFloat := 0;
    JourTabBSUMME.AsFloat := 0;

    JourTabKOST_NETTO.AsFloat := 0;
    JourTabWERT_NETTO.AsFloat := 0;
    JourTabLohn.AsFloat := 0;
    JourTabWare.AsFloat := 0;
    JourTabRohgewinn.AsFloat := 0;
    JourTabTKost.AsFloat := 0;
    JourTabPROVIS_WERT.AsFloat := 0;

    JourTabVRENUM.AsString := JourTabVRENUM.AsString + ' - STORNO -';

    JourTab.Post;
    JourTab.Close;

    //Datei-Links löschen
    LinkForm.DelLinks(VK_RECH, JOURNAL_ID);

    Result := True;
  EXCEPT
    MessageDlg(_('Fehler beim Storno der VK-Rechnung !'), mterror, [mbok], 0);
    Result := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Storno_Angebot(Journal_ID : Integer) : Boolean; // True, Wenn OK
BEGIN
  //     Result :=False;
  TRY
    JourTab.Close;
    JourTab.ParamByName('ID').Value := Journal_ID;
    JourTab.Open;

    JPosTab.Close;
    JPosTab.ParamByName('ID').Value := Journal_ID;
    JPosTab.Open;

    WHILE NOT JPosTab.Eof DO
      JPosTab.Delete;
    JPosTab.Close;

    JourTab.Delete;
    JourTab.Close;

    //Datei-Links löschen
    LinkForm.DelLinks(VK_AGB, JOURNAL_ID);

    Result := True;
  EXCEPT
    MessageDlg(_('Fehler beim Storno des Angebotes !'), mterror, [mbok], 0);
    Result := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Storno_EKBestellung(Journal_ID : Integer) : Boolean; // True, Wenn OK
BEGIN
  //     Result :=False;
  TRY
    JPosTab.Close;
    JPosTab.ParamByName('ID').Value := Journal_ID;
    JPosTab.Open;

    WHILE NOT JPosTab.Eof DO
      BEGIN
        // Bestellmenge aktualsieren
        IF (JPosTabArtikelTyp.Value = 'N') AND
          (JPosTabARTIKEL_ID.Value > -1) THEN
          BEGIN
            // Menge nicht verändern !!!
            ArtMengeTab.Close;
            ArtMengeTab.ParamByName('ID').Value := JPosTabARTIKEL_ID.Value;
            ArtMengeTab.ParamByName('SUBMENGE').Value := 0;
            //Bestellmenge erniedrigen
            //ArtMengeTab.ParamByName ('BMENGE').Value :=JPosTabMenge.Value; // Menge wird abgezogen
            ArtMengeTab.ExecSql;
          END;

        // Position löschen
        JPosTab.Delete;
      END;
    JPosTab.Close;

    JourTab.Close;
    JourTab.ParamByName('ID').AsInteger := Journal_ID;
    JourTab.Open;
    JourTab.Delete;
    JourTab.Close;

    //Datei-Links löschen
    LinkForm.DelLinks(EK_BEST, JOURNAL_ID);

    Result := True;
  EXCEPT
    MessageDlg(_('Fehler beim Storno der EK-Bestellung !'), mterror, [mbok], 0);
    Result := False;
  END;

  UpdateEKBestMenge; // Bestellmengen aktualisieren
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.Storno_Lieferschein(Journal_ID : Integer) : Boolean; // True, Wenn OK
VAR
  S : STRING;
BEGIN
  //     Result :=False;
  TRY
    JourTab.Close;
    JourTab.ParamByName('ID').Value := Journal_ID;
    JourTab.Open;

    // ggf. Seriennummer freigeben
    WITH UniQuery DO
      BEGIN
        Close;
        //SQL.Clear;
        SQL.Text := 'SELECT * FROM JOURNALPOS_SERNUM ' +
          'WHERE JOURNAL_ID=' + InttoStr(Journal_ID);
        Open;
        S := '';
        IF RecordCount > 0 THEN
          BEGIN
            IF length(s) > 0 THEN
              s := s + ',';
            s := s + IntToStr(FieldByName('SNUM_ID').AsInteger);
            Next;
          END;
        Close;

        IF length(s) > 0 THEN
          BEGIN
            SQL.Text := 'UPDATE ARTIKEL_SERNUM SET STATUS="LAGER" ' +
              'WHERE SNUM_ID IN (' + s + ')';
            ExecSql;
          END;
      END; //

    JPosTab.Close;
    JPosTab.ParamByName('ID').Value := Journal_ID;
    JPosTab.Open;

    WHILE NOT JPosTab.Eof DO
      JPosTab.Delete;
    JPosTab.Close;

    JourTab.Delete;
    JourTab.Close;

    //Datei-Links löschen
    LinkForm.DelLinks(VK_LIEF, JOURNAL_ID);

    Result := True;
  EXCEPT
    MessageDlg(_('Fehler beim Storno des Lieferscheins !'), mterror, [mbok], 0);
    Result := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.CalcLeitWaehrung(Betrag : Double; Waehrung : STRING) : Double;
VAR
  Kurs : Double;
BEGIN
  IF Waehrung = Leitwaehrung THEN
    BEGIN
      Result := Betrag;
      exit;
    END;
  IF CacheLastWaehrung <> Waehrung THEN
    BEGIN
      Kurs := ReadDouble('MAIN\WAEHRUNG', Waehrung, 1);
      CacheLastKurs := Kurs;
      CacheLastWaehrung := Waehrung;
    END
  ELSE
    Kurs := CacheLastKurs;

  IF Kurs = 0 THEN
    Result := 0
  ELSE
    Result := (CAO_Round(Betrag / Kurs * 100)) / 100;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.CopyRechnung(Journal_ID, Dest : Integer) : Integer; // Liefert Rec-ID zurück
VAR
  id, i : integer;
  BelegNr : STRING;
  w : STRING;
  N, n0, n1, n2, n3,
    M, m0, m1, m2, m3,
    B, b0, b1, b2, b3 : Double;

  //    Summe,
  //    NSumme,MSumme,BSumme, Steuer,
  EK, Lohn, Ware, TKst : Double;
  T : Char;

BEGIN
  Result := -1;

  N := 0;
  N0 := 0;
  N1 := 0;
  N2 := 0;
  N3 := 0;
  M := 0;
  M0 := 0;
  M1 := 0;
  M2 := 0;
  M3 := 0;
  B := 0;
  B0 := 0;
  B1 := 0;
  B2 := 0;
  B3 := 0;
  Lohn := 0;
  Ware := 0;
  TKst := 0;

  CpySrcKopfTab.Close;
  CpySrcKopfTab.ParamByName('ID').Value := Journal_ID;
  CpySrcKopfTab.Open;

  IF CpySrcKopfTab.RecordCount = 0 THEN
    BEGIN
      CpySrcKopfTab.Close;
      Exit;
    END;

  BelegNr := IncNummerStr(Dest);

  CpyDstKopfTab.Open;
  CpyDstKopfTab.Append;
  IF CpySrcKopfTabWAEHRUNG.Value = '' THEN
    CpySrcKopfTabWAEHRUNG.Value := ReadString('MAIN', 'LEITWAEHRUNG', '€');
  w := CpySrcKopfTabWAEHRUNG.Value;

  CpyDstKopfTabQUELLE.Value := Dest; //CpySrcKopfTabQUELLE.Value;
  CpyDstKopfTabQUELLE_SUB.Value := 0; //CpySrcKopfTabQUELLE_SUB.Value;

  CpyDstKopfTabATRNUM.Value := -1;
  CpyDstKopfTabLief_Addr_ID.Value := CpySrcKopfTabLief_Addr_ID.Value;

  CpyDstKopfTabVRENUM.Value := BelegNr;
  CpyDstKopfTabVLSNUM.Value := '';
  CpyDstKopfTabFOLGENR.Value := -1;
  CpyDstKopfTabKM_STAND.Value := -1;
  CpyDstKopfTabADATUM.Value := 0;
  CpyDstKopfTabRDATUM.Value := Now;

  IF Dest <> EK_BEST_EDI THEN
    CpyDstKopfTabLDATUM.Value := 0
  ELSE
    CpyDstKopfTabLDATUM.Value := CpySrcKopfTabLDATUM.Value;

  CpyDstKopfTabTermin.Value := 0;

  CpyDstKopfTabKOST_NETTO.Value := Calcleitwaehrung(CpySrcKopfTabKOST_NETTO.AsFloat, w);
  CpyDstKopfTabWERT_NETTO.Value := CalcLeitWaehrung(CpySrcKopfTabWERT_NETTO.AsFloat, W);
  CpyDstKopfTabLOHN.Value := CalcLeitWaehrung(CpySrcKopfTabLOHN.AsFloat, W);
  CpyDstKopfTabWARE.Value := CalcLeitWaehrung(CpySrcKopfTabWARE.AsFloat, W);
  CpyDstKopfTabTKOST.Value := CalcLeitWaehrung(CpySrcKopfTabTKOST.AsFloat, W);

  CpyDstKopfTabADDR_ID.Value := CpySrcKopfTabADDR_ID.Value;
  CpyDstKopfTabKFZ_ID.Value := CpySrcKopfTabKFZ_ID.Value;
  CpyDstKopfTabVERTRETER_ID.Value := CpySrcKopfTabVERTRETER_ID.Value;
  CpyDstKopfTabGLOBRABATT.Value := CpySrcKopfTabGLOBRABATT.Value;

  CpyDstKopfTabPR_EBENE.Value := CpySrcKopfTabPR_EBENE.Value;
  CpyDstKopfTabLIEFART.Value := CpySrcKopfTabLIEFART.Value;
  CpyDstKopfTabZAHLART.Value := CpySrcKopfTabZAHLART.Value;

  CpyDstKopfTabMWST_0.Value := CpySrcKopfTabMWST_0.Value;
  CpyDstKopfTabMWST_1.Value := CpySrcKopfTabMWST_1.Value;
  CpyDstKopfTabMWST_2.Value := CpySrcKopfTabMWST_2.Value;
  CpyDstKopfTabMWST_3.Value := CpySrcKopfTabMWST_3.Value;
  CpyDstKopfTabMWST_FREI_FLAG.Value := CpySrcKopfTabMWST_FREI_FLAG.Value;

  CpyDstKopfTabNSUMME_0.Value := CalcLeitWaehrung(CpySrcKopfTabNSUMME_0.AsFloat, W);
  CpyDstKopfTabNSUMME_1.Value := CalcLeitWaehrung(CpySrcKopfTabNSUMME_1.AsFloat, W);
  CpyDstKopfTabNSUMME_2.Value := CalcLeitWaehrung(CpySrcKopfTabNSUMME_2.AsFloat, W);
  CpyDstKopfTabNSUMME_3.Value := CalcLeitWaehrung(CpySrcKopfTabNSUMME_3.AsFloat, W);
  CpyDstKopfTabNSUMME.Value := CalcLeitWaehrung(CpySrcKopfTabNSUMME.AsFloat, W);

  CpyDstKopfTabMSUMME_0.Value := CalcLeitWaehrung(CpySrcKopfTabMSUMME_0.AsFloat, W);
  CpyDstKopfTabMSUMME_1.Value := CalcLeitWaehrung(CpySrcKopfTabMSUMME_1.AsFloat, W);
  CpyDstKopfTabMSUMME_2.Value := CalcLeitWaehrung(CpySrcKopfTabMSUMME_2.AsFloat, W);
  CpyDstKopfTabMSUMME_3.Value := CalcLeitWaehrung(CpySrcKopfTabMSUMME_3.AsFloat, W);
  CpyDstKopfTabMSUMME.Value := CalcLeitWaehrung(CpySrcKopfTabMSUMME.AsFloat, W);

  CpyDstKopfTabBSUMME_0.Value := CalcLeitWaehrung(CpySrcKopfTabBSUMME_0.AsFloat, W);
  CpyDstKopfTabBSUMME_1.Value := CalcLeitWaehrung(CpySrcKopfTabBSUMME_1.AsFloat, W);
  CpyDstKopfTabBSUMME_2.Value := CalcLeitWaehrung(CpySrcKopfTabBSUMME_2.AsFloat, W);
  CpyDstKopfTabBSUMME_3.Value := CalcLeitWaehrung(CpySrcKopfTabBSUMME_3.AsFloat, W);
  CpyDstKopfTabBSUMME.Value := CalcLeitWaehrung(CpySrcKopfTabBSUMME.AsFloat, W);

  CpyDstKopfTabATSUMME.Value := CalcLeitWaehrung(CpySrcKopfTabATSUMME.AsFloat, W);
  CpyDstKopfTabATMSUMME.Value := CalcLeitWaehrung(CpySrcKopfTabATMSUMME.AsFloat, W);

  CpyDstKopfTabWAEHRUNG.Value := LeitWaehrung; //CpySrcKopfTabWAEHRUNG.Value;
  CpyDstKopfTabGEGENKONTO.Value := CpySrcKopfTabGEGENKONTO.Value;
  CpyDstKopfTabSOLL_STAGE.Value := CpySrcKopfTabSOLL_STAGE.Value;
  CpyDstKopfTabSOLL_SKONTO.Value := CpySrcKopfTabSOLL_SKONTO.Value;
  CpyDstKopfTabSOLL_NTAGE.Value := CpySrcKopfTabSOLL_NTAGE.Value;
  CpyDstKopfTabSOLL_RATEN.Value := CpySrcKopfTabSOLL_RATEN.Value;

  CpyDstKopfTabSOLL_RATBETR.Value := CalcLeitWaehrung(CpySrcKopfTabSOLL_RATBETR.AsFloat, W);

  CpyDstKopfTabSOLL_RATINTERVALL.Value := CpySrcKopfTabSOLL_RATINTERVALL.Value;

  CpyDstKopfTabIST_ANZAHLUNG.Value := 0;
  CpyDstKopfTabIST_ZAHLDAT.Value := 0;
  CpyDstKopfTabMAHNKOSTEN.Value := 0;
  CpyDstKopfTabKONTOAUSZUG.Value := -1;
  CpyDstKopfTabBANK_ID.Value := -1;
  CpyDstKopfTabSTADIUM.Value := 6;
  CpyDstKopfTabFREIGABE1_Flag.Value := False;
  CpyDstKopfTabERSTELLT.Value := Now;
  CpyDstKopfTabERST_NAME.Value := View_User;

  CpyDstKopfTabKUN_NUM.Value := CpySrcKopfTabKUN_NUM.Value;
  CpyDstKopfTabKUN_ANREDE.Value := CpySrcKopfTabKUN_ANREDE.Value;
  CpyDstKopfTabKUN_NAME1.Value := CpySrcKopfTabKUN_NAME1.Value;
  CpyDstKopfTabKUN_NAME2.Value := CpySrcKopfTabKUN_NAME2.Value;
  CpyDstKopfTabKUN_NAME3.Value := CpySrcKopfTabKUN_NAME3.Value;
  CpyDstKopfTabKUN_ABTEILUNG.Value := CpySrcKopfTabKUN_ABTEILUNG.Value;
  CpyDstKopfTabKUN_STRASSE.Value := CpySrcKopfTabKUN_STRASSE.Value;
  CpyDstKopfTabKUN_LAND.Value := CpySrcKopfTabKUN_LAND.Value;
  CpyDstKopfTabKUN_PLZ.Value := CpySrcKopfTabKUN_PLZ.Value;
  CpyDstKopfTabKUN_ORT.Value := CpySrcKopfTabKUN_ORT.Value;
  CpyDstKopfTabUSR1.Value := CpySrcKopfTabUSR1.Value;
  CpyDstKopfTabUSR2.Value := CpySrcKopfTabUSR2.Value;
  CpyDstKopfTabPROJEKT.Value := CpySrcKopfTabPROJEKT.Value;
  CpyDstKopfTabORGNUM.Value := CpySrcKopfTabORGNUM.Value;
  CpyDstKopfTabBEST_NAME.Value := CpySrcKopfTabBEST_NAME.Value;
  CpyDstKopfTabBEST_CODE.Value := CpySrcKopfTabBEST_CODE.Value;
  CpyDstKopfTabINFO.AsString := CpySrcKopfTabINFO.AsString;
  CpyDstKopfTabBEST_DATUM.Value := CpySrcKopfTabBEST_DATUM.Value;
  CpyDstKopfTabPROVIS_WERT.Value := CpySrcKopfTabPROVIS_WERT.Value;

  CpyDstKopfTabPROVIS_WERT.Value := CpySrcKopfTabPROVIS_WERT.Value;
  CpyDstKopfTabBRUTTO_FLAG.AsBoolean := CpySrcKopfTabBRUTTO_FLAG.AsBoolean;

  CpyDstKopfTabGewicht.AsFloat := CpySrcKopfTabGewicht.AsFloat;
  CpyDstKopfTabRohgewinn.AsFloat := CpySrcKopfTabRohgewinn.AsFloat;

  CpyDstKopfTabUW_NUM.Value := -1;

  IF ((CpySrcKopfTabQuelle.AsInteger IN [EK_RECH, EK_RECH_EDI]) AND (Dest =
    VK_RECH_EDI)) OR
    ((CpySrcKopfTabQuelle.AsInteger IN [VK_RECH, VK_RECH_EDI, VK_AGB,
    VK_AGB_EDI]) AND (Dest = EK_BEST_EDI)) THEN
    BEGIN
      CpyDstKopfTabKOST_NETTO.Value := 0;
      CpyDstKopfTabWERT_NETTO.Value := 0;
      CpyDstKopfTabLOHN.Value := 0;
      CpyDstKopfTabWARE.Value := 0;
      CpyDstKopfTabTKOST.Value := 0;
      CpyDstKopfTabGLOBRABATT.Value := 0;

      IF Dest <> EK_BEST_EDI THEN
        BEGIN
          CpyDstKopfTabPR_EBENE.Value := 2;
        END
      ELSE
        BEGIN
          CpyDstKopfTabPR_EBENE.Value := 0;
          CpyDstKopfTabKUN_NUM.Value := '';
          CpyDstKopfTabKUN_ANREDE.Value := '';
          CpyDstKopfTabKUN_NAME1.Value := '';
          CpyDstKopfTabKUN_NAME2.Value := '';
          CpyDstKopfTabKUN_NAME3.Value := '';
          CpyDstKopfTabKUN_ABTEILUNG.Value := '';
          CpyDstKopfTabKUN_STRASSE.Value := '';
          CpyDstKopfTabKUN_LAND.Value := '';
          CpyDstKopfTabKUN_PLZ.Value := '';
          CpyDstKopfTabKUN_ORT.Value := '';

          CpyDstKopfTabLief_Addr_ID.Value := CpySrcKopfTabAddr_ID.Value;
          CpyDstKopfTabAddr_ID.Value := -1;
          CpyDstKopfTabLDatum.Value := int(Now) + 3;
        END;

      CpyDstKopfTabLIEFART.Value := -1;
      CpyDstKopfTabZAHLART.Value := -1;
    END;

  CpyDstKopfTab.Post;

  ID := CpyDstKopfTabREC_ID.Value;

  CpySrcPosTab.Close;
  CpySrcPosTab.ParamByName('ID').AsInteger := Journal_ID;
  CpySrcPosTab.Open;

  CpyDstPosTab.Open;

  WHILE NOT CpySrcPosTab.Eof DO
    BEGIN
      // Bei Lieferscheinen die Transportkosten überspringen
      IF (Dest = VK_LIEF_EDI) AND
        (CpySrcPosTab.FieldByName('ARTIKELTYP').AsString = 'K') THEN
        BEGIN
          CpySrcPosTab.Next;
          Continue;
        END;

      CpyDstPosTab.Append;

      FOR i := 0 TO CpySrcPosTab.Fields.Count - 1 DO
        BEGIN
          IF uppercase(CpySrcPosTab.Fields[i].FieldName) <> 'REC_ID' THEN
            CpyDstPosTab.FieldByName(CpySrcPosTab.Fields[i].FieldName).Value :=
              CpySrcPosTab.Fields[i].Value;
        END;
      CpyDstPosTab.FieldByName('VRENUM').Value := BelegNr;
      CpyDstPosTab.FieldByName('QUELLE').Value := Dest;
      CpyDstPosTab.FieldByName('QUELLE_SUB').Value := 0;
      //CpyDstPosTab.FieldByName('JAHR').Value            :=0;
      CpyDstPosTab.FieldByName('JOURNAL_ID').Value := ID;
      CpyDstPosTab.FieldByName('GEBUCHT').Value := False;
      CpyDstPosTab.FieldByName('EPREIS').Value :=
        CalcLeitwaehrung(CpySrcPosTab.FieldByName('EPREIS').AsFloat, W);
      CpyDstPosTab.FieldByName('E_RGEWINN').Value :=
        CalcLeitwaehrung(CpySrcPosTab.FieldByName('E_RGEWINN').AsFloat, W);

      //NEU
      //if (CpySrcKopfTabQuelle.AsInteger in [EK_RECH,EK_RECH_EDI])and(Dest=VK_RECH_EDI) then
      IF ((CpySrcKopfTabQuelle.AsInteger IN [EK_RECH, EK_RECH_EDI, VK_Auftrag])
        AND (Dest = VK_RECH_EDI)) OR
        ((CpySrcKopfTabQuelle.AsInteger IN [VK_RECH, VK_RECH_EDI, VK_AGB,
        VK_AGB_EDI]) AND (Dest = EK_BEST_EDI)) OR
        (Dest = VK_LIEF_EDI) THEN
        BEGIN
          //         showmessage('drinn in '+inttostr(dest));
          IF (DEST <> EK_BEST_EDI) THEN
            CpyDstPosTab.FieldByName('GEGENKTO').Value := ReadInteger('MAIN\KONTEN', 'DEF_ERLOESKTO', 8400)
          ELSE
            CpyDstPosTab.FieldByName('GEGENKTO').Value := ReadInteger('MAIN\KONTEN', 'DEF_AUFWANDSKTO', 3400);

          IF (CpyDstPosTab.FieldByName('ARTIKEL_ID').AsInteger >= 0) AND
            (Dest <> VK_LIEF_EDI) AND (CpySrcKopfTabQuelle.Value <> VK_Auftrag)
            THEN
            BEGIN
              UpdateArtTab.Close;
              UpdateArtTab.ParamByName('ID').AsInteger := CpyDstPosTab.FieldByName('ARTIKEL_ID').AsInteger;
              UpdateArtTab.Open;

              IF UpdateArtTab.RecordCount = 1 THEN
                BEGIN
                  IF (DEST = EK_BEST_EDI) AND // nur bei EK-Bestellungen
                    (UpdateArtTab.FieldByName('RABGRP_ID').AsString <> '-') AND
                    (length(UpdateArtTab.FieldByName('RABGRP_ID').AsString) > 0)
                    THEN
                    BEGIN
                      // Ist ein Artikel mit Rabattgruppe !!!
                      CASE AnzPreis OF
                        1 : EK := UpdateArtTab.FieldByName('VK1').Value;
                        2 : EK := UpdateArtTab.FieldByName('VK2').Value;
                        3 : EK := UpdateArtTab.FieldByName('VK3').Value;
                        4 : EK := UpdateArtTab.FieldByName('VK4').Value;
                        ELSE
                          EK := UpdateArtTab.FieldByName('VK5').Value;
                      END;

                      CalcRabGrpPreis(UpdateArtTab.FieldByName('RABGRP_ID').AsString, 0 {EK}, EK);
                      CpyDstPosTab.FieldByName('EPREIS').Value := CAO_round(EK *
                        1000) / 1000; //auf 3 Nachkommastellen runden
                    END
                  ELSE
                    BEGIN
                      // neuen VK-Preis zuweisen
                      CASE CpyDstKopfTabPR_EBENE.Value OF
                        0 : CpyDstPosTab.FieldByName('EPREIS').Value := UpdateArtTab.FieldByName('EK_PREIS').Value;
                        1 : CpyDstPosTab.FieldByName('EPREIS').Value := UpdateArtTab.FieldByName('VK1').Value;
                        2 : CpyDstPosTab.FieldByName('EPREIS').Value := UpdateArtTab.FieldByName('VK2').Value;
                        3 : CpyDstPosTab.FieldByName('EPREIS').Value := UpdateArtTab.FieldByName('VK3').Value;
                        4 : CpyDstPosTab.FieldByName('EPREIS').Value := UpdateArtTab.FieldByName('VK4').Value;
                        ELSE
                          CpyDstPosTab.FieldByName('EPREIS').Value := UpdateArtTab.FieldByName('VK5').Value;
                      END;
                    END;
                END;

              IF DEST <> EK_BEST_EDI THEN
                CpyDstPosTab.FieldByName('GEGENKTO').Value := UpdateArtTab.FieldByName('ERLOES_KTO').AsInteger
              ELSE
                CpyDstPosTab.FieldByName('GEGENKTO').Value := UpdateArtTab.FieldByName('AUFW_KTO').AsInteger;

              IF DEST = EK_BEST_EDI THEN
                CpyDstPosTab.FieldByName('RABATT').Value := 0;
              IF DEST = EK_BEST_EDI THEN
                CpyDstPosTab.FieldByName('RABATT2').Value := 0;
              IF DEST = EK_BEST_EDI THEN
                CpyDstPosTab.FieldByName('RABATT3').Value := 0;

              UpdateArtTab.Close;
            END;

          {
          Summe :=CpyDstPosTab.FieldByName('EPREIS').Value * CpyDstPosTab.FieldByName('MENGE').Value;
          if CpyDstPosTab.FieldByName('RABATT').Value <> 0
           then Summe :=Summe - Summe * CpyDstPosTab.FieldByName('RABATT').Value / 100;

          Case CpyDstPosTab.FieldByName('STEUER_CODE').Value of
             0: Steuer :=CpyDstKopfTabMWST_0.Value;
             1: Steuer :=CpyDstKopfTabMWST_1.Value;
             2: Steuer :=CpyDstKopfTabMWST_2.Value;
             3: Steuer :=CpyDstKopfTabMWST_3.Value;
             else Steuer :=0;
          end;

          NSumme :=CAO_round_nk(Summe,2);  // Auf ganze Pfennige Runden
          MSumme :=CAO_round_nk(Summe * (Steuer / 100),2); // Auf ganze Pfennige Runden
          BSumme :=NSumme+MSumme;

          N :=N+NSumme;
          M :=M+MSumme;
          B :=B+BSumme;

          case CpyDstPosTab.FieldByName('STEUER_CODE').Value of
                 0:M0 :=M0 + MSumme;
                 1:M1 :=M1 + MSumme;
                 2:M2 :=M2 + MSumme;
                 3:M3 :=M3 + MSumme;
          end;
          }

          IF CpyDstKopfTabBrutto_Flag.ASBoolean THEN
            BEGIN
              //Bruttofakturierung
              B := B + CpyDstPosTab.FieldByName('GPREIS').AsFloat;

              CASE CpyDstPosTab.FieldByName('STEUER_CODE').AsInteger OF
                0 : B0 := B0 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
                1 : B1 := B1 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
                2 : B2 := B2 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
                3 : B3 := B3 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
              END;
            END
          ELSE
            BEGIN
              // Netto
              N := N + CpyDstPosTab.FieldByName('GPREIS').AsFloat;

              CASE CpyDstPosTab.FieldByName('STEUER_CODE').AsInteger OF
                0 : N0 := N0 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
                1 : N1 := N1 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
                2 : N2 := N2 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
                3 : N3 := N3 + CpyDstPosTab.FieldByName('GPREIS').AsFloat;
              END;
            END;

          // Lohn, Ware, Transportkosten
          IF length(CpyDstPosTab.FieldByName('ARTIKELTYP').AsString) = 1 THEN
            T := CpyDstPosTab.FieldByName('ARTIKELTYP').AsString[1]
          ELSE
            T := '?';

          CASE t OF
            'N', 'S', 'V', 'F' : Ware := Ware +
              CpyDstPosTab.FieldByName('GPREIS').AsFloat; //NSumme;
            'L' : Lohn := Lohn + CpyDstPosTab.FieldByName('GPREIS').AsFloat; //NSumme;
            'K' : TKst := TKst + CpyDstPosTab.FieldByName('GPREIS').AsFloat; //NSumme;
          END;
        END;

      IF NOT CpyDstKopfTabMWST_FREI_Flag.AsBoolean THEN
        BEGIN
          IF CpyDstKopfTabBrutto_Flag.ASBoolean THEN
            BEGIN
              //Bruttofakturierung
              M1 := CAO_round_nk(B1 / (100 + CpyDstKopfTabMwSt_1.AsFloat) * CpyDstKopfTabMwSt_1.AsFloat, 2);
              M2 := CAO_round_nk(B2 / (100 + CpyDstKopfTabMwSt_2.AsFloat) * CpyDstKopfTabMwSt_2.AsFloat, 2);
              M3 := CAO_round_nk(B3 / (100 + CpyDstKopfTabMwSt_3.AsFloat) * CpyDstKopfTabMwSt_3.AsFloat, 2);

              N0 := B0 - M0;
              N1 := B1 - M1;
              N2 := B2 - M2;
              N3 := B3 - M3;

              M := M0 + M1 + M2 + M3;
              N := B - M;
            END
          ELSE
            BEGIN
              // Nettofakturierung
              M1 := CAO_round(N1 * CpyDstKopfTabMwSt_1.AsFloat) / 100;
              M2 := CAO_round(N2 * CpyDstKopfTabMwSt_2.AsFloat) / 100;
              M3 := CAO_round(N3 * CpyDstKopfTabMwSt_3.AsFloat) / 100;

              B0 := N0 + M0;
              B1 := N1 + M1;
              B2 := N2 + M2;
              B3 := N3 + M3;

              M := M0 + M1 + M2 + M3;
              B := N + M;
            END;
        END
      ELSE
        BEGIN // MwSt-Frei ...
          M := 0;
          M1 := 0;
          M2 := 0;
          M3 := 0;
          B := N;
        END;

      // Bei VK-Rechnungen (Edi) die Lieferscheinnummer der
      // zugrundeliegenden Rechnung löschen
      IF Dest = VK_RECH_EDI THEN
        CpyDstPosTab.FieldByName('VLSNUM').AsString := '';

      IF ((Dest = EK_RECH_EDI) AND (CpySrcKopfTabQuelle.AsInteger = EK_BEST)) OR
        ((Dest = VK_LIEF_EDI) AND (CpySrcKopfTabQuelle.AsInteger = VK_RECH_EDI))
        THEN
        BEGIN
          // Im Einkauf Verweis auf die zugrundeliegenden Bestell-Positionen machen
          // im Verkauf/Lieferschein Verweis auf zugrundeliegende EDI-Rechnunsposition machen
          CpyDstPosTab.FieldByName('QUELLE_SRC').AsInteger :=
            CpySrcPosTab.FieldByName('REC_ID').AsInteger;
        END;

      CpyDstPosTab.Post;

      // Seriennummern bei Lieferscheinen zuweisen
      {
      if (Dest=VK_LIEF_EDI) and
         (CpySrcKopfTabQuelle.AsInteger=VK_RECH_EDI) then
      begin
          UniQuery.Sql.Text :=
            'UPDATE ARTIKEL_SERNUM SET '+
            'LS_JOURNALPOS_ID='+
            Inttostr(CpyDstPosTab.FieldByName ('REC_ID').AsInteger)+','+
            'LS_JOURNAL_ID='+Inttostr(ID)+
            ' where VK_JOURNALPOS_ID='+
            IntToStr(CpySrcPosTab.FieldByName ('REC_ID').AsInteger)+' and '+
            'LS_JOURNALPOS_ID=-1 and ARTIKEL_ID='+
            IntToStr(CpyDstPosTab.FieldByName('ARTIKEL_ID').AsInteger);

          UniQuery.ExecSql;
      end;
      }

      // EDI-Mengen aktualisieren
      UpdateArtikelEdiMenge(Dest,
        CpyDstPosTab.FieldByName('ARTIKEL_ID').AsInteger,
        0);

      CpySrcPosTab.Next;
    END;

  IF ((CpySrcKopfTabQuelle.AsInteger IN [EK_RECH, EK_RECH_EDI]) AND (Dest =
    VK_RECH_EDI)) OR
    ((CpySrcKopfTabQuelle.AsInteger IN [VK_RECH, VK_RECH_EDI, VK_AGB,
    VK_AGB_EDI]) AND (Dest = EK_BEST_EDI)) OR
    (Dest = VK_LIEF_EDI) THEN
    BEGIN
      CpyDstKopfTab.Edit;

      CpyDstKopfTabNSumme.Value := N;
      CpyDstKopfTabNSumme_0.Value := N0;
      CpyDstKopfTabNSumme_1.Value := N1;
      CpyDstKopfTabNSumme_2.Value := N2;
      CpyDstKopfTabNSumme_3.Value := N3;

      CpyDstKopfTabMSumme.Value := M;
      CpyDstKopfTabMSumme_0.Value := M0;
      CpyDstKopfTabMSumme_1.Value := M1;
      CpyDstKopfTabMSumme_2.Value := M2;
      CpyDstKopfTabMSumme_3.Value := M3;

      CpyDstKopfTabBSumme.Value := B;
      CpyDstKopfTabBSumme_0.Value := B0;
      CpyDstKopfTabBSumme_1.Value := B1;
      CpyDstKopfTabBSumme_2.Value := B2;
      CpyDstKopfTabBSumme_3.Value := B3;

      CpyDstKopfTabWare.Value := Ware;
      CpyDstKopfTabLohn.Value := Lohn;
      CpyDstKopfTabTKost.Value := TKst;

      CpyDstKopfTab.Post;
    END;
  CpySrcPosTab.Close;
  CpyDstPosTab.Close;
  CpyDstKopfTab.Close;
  CpySrcKopfTab.Close;

  Result := ID;
END;

//------------------------------------------------------------------------------

FUNCTION tDM1.BucheKasse(Datum : tDateTime;
  Quelle : Integer;
  Journal_ID : Integer;
  BelNum : STRING;
  GKonto : Integer;
  Skonto : Double;
  Betrag : Double;
  Text : STRING) : Boolean;

BEGIN
  KasBuch.Close;
  KasBuch.Open;
  KasBuch.Append;

  TRY
    KasBuchBDATUM.Value := Datum;
    KasBuchQUELLE.Value := Quelle;
    KasBuchJOURNAL_ID.Value := Journal_ID;
    KasBuchZU_ABGANG.Value := Betrag;
    KasBuchBTXT.AsString := Text;
    KasBuchBELEGNUM.Value := BelNum;
    KasBuchGKONTO.Value := GKonto;
    KasBuchSKONTO.Value := Skonto;

    KasBuchERSTELLT.AsDateTime := now;
    KasBuchERST_NAME.AsString := view_user;
    KasBuchMA_ID.AsInteger := MitarbeiterID;

    KasBuch.Post;
    Result := True;
  EXCEPT
    KasBuch.Cancel;
    Result := False;
  END;
  KasBuch.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.JourTabCalcFields(DataSet : TDataSet);
BEGIN
  JourTabIST_SKONTO_BETR.Value := JourTabBSumme.Value - JourTabIST_Betrag.Value;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.KGRTabBeforePost(DataSet : TDataSet);
BEGIN
  IF KGRTabMainKey.Value = '' THEN
    KGRTabMainKey.Value := 'MAIN\ADDR_HIR';
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.FirBankTabBeforePost(DataSet : TDataSet);
BEGIN
  FirBankTabmainkey.Value := 'MAIN\FIRMENKONTEN';
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.GetWGRDefaultKonten(WGR : Integer; VAR EKTO, AKTO : Integer) :
  Boolean;
BEGIN
  Result := False;
  IF NOT WGRTab.Active THEN
    WGRTab.Open;

  IF WGRTab.Locate('ID', WGR, []) THEN
    BEGIN
      IF WgrTabDEF_AKTO.AsInteger > 0 THEN
        AKTO := WgrTabDEF_AKTO.AsInteger;
      IF WGRTabDEF_EKTO.AsInteger > 0 THEN
        EKTO := WGRTabDEF_EKTO.AsInteger;
      Result := True;
    END;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.CheckBankverbindung(addr_id : integer) : boolean;
BEGIN
  //     Result :=False;
  KunTab.Close;
  KunTab.ParamByName('ID').AsInteger := addr_id;
  KunTab.Open;
  IF NOT KunTab.Eof THEN
    BEGIN
      TRY
        Result := (length(KunTabBLZ.AsString) = BLZ_LEN) AND
          (length(KunTabKTO.AsString) > 1) AND
          (length(KunTabBank.AsString) > 3) AND
          (StrToInt(KunTabBLZ.AsString) > 0);

      EXCEPT
        Result := False;
      END;
    END
  ELSE
    Result := False;
  KunTab.Close;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.GetBankverbindung(addr_id : integer;
  VAR BLZ : Integer;
  VAR KTO : STRING;
  VAR Inhaber : STRING) : boolean;
VAR
  OK : Boolean;
BEGIN
  OK := False;
  KunTab.Close;
  KunTab.ParamByName('ID').AsInteger := addr_id;
  KunTab.Open;
  IF NOT KunTab.Eof THEN
    BEGIN
      TRY
        OK := (length(KunTabBLZ.AsString) = 8) AND
          (length(KunTabKTO.AsString) > 1) AND
          (length(KunTabBank.AsString) > 3) AND
          (StrToInt(KunTabBLZ.AsString) > 0);
      EXCEPT
        OK := False;
      END;
    END
  ELSE
    Result := OK;

  IF OK THEN
    BEGIN
      BLZ := StrToInt(KunTabBLZ.AsString);
      KTO := KunTabKTO.AsString;

      IF Length(KunTabKTO_INHABER.AsString) > 0 THEN
        Inhaber := KunTabKTO_INHABER.AsString
      ELSE
        Inhaber := KunTabName1.AsString;
    END;

  KunTab.Close;
  Result := OK;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.GetLieferant(addr_id : integer; VAR Info : STRING) : boolean; // True wenn ok
VAR
  S : STRING;
BEGIN
  Result := FALSE;

  IF Addr_id < 1 THEN
    BEGIN
      Info := '???';
      exit;
    END;

  KunTab.Close;
  KunTab.ParamByName('ID').AsInteger := addr_id;
  KunTab.Open;
  IF NOT KunTab.Eof THEN
    BEGIN
      S := KunTabNAME1.AsString;
      IF length(KunTabNAME2.AsString) > 0 THEN
        BEGIN
          IF length(S) > 0 THEN
            S := S + ', ';
          S := S + KunTabNAME2.AsString;
        END;
      IF length(KunTabNAME3.AsString) > 0 THEN
        BEGIN
          IF length(S) > 0 THEN
            S := S + ', ';
          S := S + KunTabNAME3.AsString;
        END;

      IF length(KunTabStrasse.AsString) > 0 THEN
        BEGIN
          IF length(S) > 0 THEN
            S := S + ', ';
          S := S + KunTabStrasse.AsString;
        END;

      S := S + ', ' + KunTabPlz.AsString + ' ' + KunTabOrt.AsString;

      Info := S;
      Result := True;
    END
  ELSE
    BEGIN
      Result := False;
      S := '';
    END;
  KunTab.Close;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.DataModuleDestroy(Sender : TObject);
BEGIN
  SetLength(MandantTab, 0);
END;

//------------------------------------------------------------------------------

FUNCTION TDM1.CalcRabGrpPreis(RGID : STRING; PR_Ebene : Integer; VAR Preis : Double) : Boolean;
BEGIN
  Result := False;
  IF (RGID = '-') OR (length(RGID) = 0) OR
    (PR_Ebene < 0) OR (PR_EBENE > 4) OR
    (PR_EBENE >= ANZPREIS) THEN
    exit;

  IF Preis = 0 THEN
    BEGIN
      Result := True;
      exit;
    END;

  IF PR_Ebene = 0 THEN //EK-Preis
    BEGIN
      IF NOT LiefRabGrp.Active THEN
        LiefRabGrp.Open;

      IF (LiefRabGrp.RecordCount = 0) THEN
        exit;
      IF (LiefRabGrpRABGRP_ID.AsString = RGID) OR
        (LiefRabGrp.Locate('RABGRP_ID', RGID, [loCaseInsensitive])) THEN
        BEGIN
          IF (LiefRabGrpRABGRP_ID.Value = RGID) THEN
            BEGIN
              Preis := Preis - (Preis / 100 * LiefRabGrpRABATT1.AsFloat);
              Preis := Preis - (Preis / 100 * LiefRabGrpRABATT2.AsFloat);
              Preis := Preis - (Preis / 100 * LiefRabGrpRABATT3.AsFloat);

              // auf 3 Nachkommastellen runden
              Preis := CAO_round(Preis * 1000) / 1000;

              Result := True;
            END;
        END;
    END
  ELSE
    BEGIN
      IF NOT KunRabGrp.Active THEN
        KunRabGrp.Open;

      IF (KunRabGrp.RecordCount = 0) THEN
        exit;
      IF //KunRabGrpRABGRP_ID.AsString=RGID)or
        (KunRabGrp.Locate('RABGRP_ID,LIEF_RABGRP',
        VarArrayOf([RGID, PR_Ebene]),
        [loCaseInsensitive])) THEN
        BEGIN
          IF (KunRabGrpRABGRP_ID.Value = RGID) AND
            (KunRabGrpLIEF_RABGRP.AsInteger = PR_Ebene) THEN
            BEGIN
              Preis := Preis - (Preis / 100 * KunRabGrpRABATT1.AsFloat);
              Preis := Preis - (Preis / 100 * KunRabGrpRABATT2.AsFloat);
              Preis := Preis - (Preis / 100 * KunRabGrpRABATT3.AsFloat);

              // auf 2 Nachkommastellen runden
              Preis := CAO_round(Preis * 100) / 100;

              Result := True;
            END;
        END;
    END;
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.UpdateArtikelEdiMenge(JournalTyp, ArtikelID : Integer;
  MengeDiff : Double) : Boolean;
VAR
  OldTickCount : DWord;
BEGIN
  //     Result :=False;

       ///NEU JAN 06.06.2004
  TRY
    OldTickCount := GetTickCount;

    zbatchsql1.sql.text :=
      'delete from ARTIKEL_BDATEN where QUELLE=' + IntToStr(JournalTyp) + ';' +
      'insert into ARTIKEL_BDATEN ' +
      'select ARTIKEL_ID, QUELLE, 0, 0, SUM(MENGE) ' +
      'from JOURNALPOS ' +
      'where QUELLE=' + IntToStr(JournalTyp) + ' and ' +
      'ARTIKELTYP IN ("N","S","L","K","X") and ARTIKEL_ID>0 ' +
      'group by ARTIKEL_ID, QUELLE ' +
      'having SUM(MENGE) !=0;';

    ProgressForm.Init('EDI-Mengen aktualisieren ...');
    ZBatchSql1.ExecSql;
    ProgressForm.Stop;

    IF (LogLevel >= 5) AND (assigned(LogForm)) THEN
      logform.addlog('EDI-Mengen aktualisieren :' + FormatFloat(',#0.00', (GetTickCount - OldTickCount) / 1000) + 'Sek.' + #13#10);

    Result := True;
  EXCEPT
    Result := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.UpdateEKBestMenge : Boolean;
VAR
  OldTickCount : DWord;
BEGIN
  //     Result :=False;

  TRY
    OldTickCount := GetTickCount;

    zbatchsql1.sql.text :=
      'delete from ARTIKEL_BDATEN where QUELLE IN(26);' +
      'insert into ARTIKEL_BDATEN ' +
      'select ' +
      'JP1.ARTIKEL_ID,26,0,0,IF(SUM(JP1.MENGE)-SUM(IFNULL(JP2.MENGE,0))<0,' +
      '0,SUM(JP1.MENGE) - SUM(IFNULL(JP2.MENGE,0))) as MENGE_OFFEN ' +
      'from (JOURNALPOS JP1, JOURNAL J )' +
      'left outer join JOURNALPOS JP2 on JP2.QUELLE_SRC=JP1.REC_ID ' +
      'and JP2.QUELLE IN(5) ' +
      'where JP1.QUELLE=6 and J.QUELLE=6 and J.STADIUM IN (20,30) and ' +
      'J.REC_ID=JP1.JOURNAL_ID and JP1.ARTIKELTYP IN ("N","S","L","K") ' +
      'group by JP1.ARTIKEL_ID having MENGE_OFFEN>0;';

    ProgressForm.Init('EK-BEST-Mengen aktualisieren ...');
    ZBatchSql1.ExecSql;
    ProgressForm.Stop;

    IF (LogLevel >= 5) AND (assigned(LogForm)) THEN
      logform.addlog('EKBEST-Mengen aktualisieren :' + FormatFloat(',#0.00',
        (GetTickCount - OldTickCount) / 1000) + 'Sek.' + #13#10);

    Result := True;
  EXCEPT
    Result := False;
  END;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.NummerTabNewRecord(DataSet : TDataSet);
BEGIN
  // Neuanlage abbrechen, es sei denn die Neuanlage wird von
  // GetNummer aufgerufen, GetNummer setzt dann die Variable
  // InNewNummer auf True
  IF NOT InNewNummer THEN
    Abort;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.ShopOrderStatusTabBeforePost(DataSet : TDataSet);
VAR
  S : STRING;
BEGIN
  IF Pos('SHOP_', ShopOrderStatusTab.Sql.Text) > 0 THEN
    S := Copy(ShopOrderStatusTab.Sql.Text, Pos('SHOP_',
      ShopOrderStatusTab.Sql.Text), 6)
  ELSE
    S := 'SHOP';

  IF ShopOrderStatusTabMainKey.Value = '' THEN
    ShopOrderStatusTabMainKey.Value := S + '\ORDERSTATUS';
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.ZBatchSql1BeforeExecute(Sender : TObject);
BEGIN
  ProgressForm.Start;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.Transact1AfterBatchExec(Sender : TObject; VAR Res : Integer);
BEGIN
  ProgressForm.UpdateScreen;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ZBatchSql1AfterExecute(Sender : TObject);
BEGIN
  ProgressForm.Stop;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ExportDatasetToStream(Stream : TStream;
  Dataset : TDataset;
  Delimiter : STRING;
  Spaltennamen : Boolean = True;
  TextInHochKomma : Boolean = True;
  DosZeichenSatz : Boolean = False);
VAR
  I, J : Integer;
  Buffer, S : STRING;
  FieldDesc : tField;
  OldC : tCursor;
BEGIN
  OldC := Screen.Cursor;
  Screen.Cursor := crSqlWait;
  I := 0;
  TRY
    WITH Dataset DO
      BEGIN
        // Feldnamen als Überschrift ausgeben !!
        IF (Fields.Count > 0) AND (Stream.Position = 0) AND (Spaltennamen) THEN
          BEGIN
            Buffer := '';
            FOR J := 0 TO FieldCount - 1 DO
              BEGIN
                IF J > 0 THEN
                  Buffer := Buffer + Delimiter;
                Buffer := Buffer + Fields[J].FieldName;
              END;

            IF DosZeichenSatz THEN
              CharToOEM(PChar(Buffer), @Buffer[1]);

            Stream.Write(PChar(Buffer)^, Length(Buffer));
          END;
        First;
        WHILE NOT EOF DO
          BEGIN
            Buffer := #13#10;
            FOR J := 0 TO Fields.Count - 1 DO
              BEGIN
                FieldDesc := Fields[J];
                IF J > 0 THEN
                  Buffer := Buffer + Delimiter;
                IF NOT (FieldDesc.DataType IN [ftInteger, ftSmallInt, ftFloat,
                  ftAutoInc, ftCurrency, ftLargeInt,
                    ftBCD]) AND
                  NOT FieldDesc.IsNull THEN
                  BEGIN
                    S := FieldDesc.AsString;

                    IF TextInHochKomma THEN
                      Buffer := Buffer + AnsiQuotedStr(Uniquery.StringToSql(S),
                        '"')
                    ELSE
                      Buffer := Buffer + Uniquery.StringToSql(S);
                  END
                ELSE
                  Buffer := Buffer + FieldDesc.AsString;
              END;
            IF DosZeichenSatz THEN
              CharToOEM(PChar(Buffer), @Buffer[1]);
            Stream.Write(PChar(Buffer)^, Length(Buffer));
            inc(i);
            Next;
          END;

        IF I = 1 THEN
          ShowMessage(Format('%d Datensatz ausgegeben.', [I]))
        ELSE
          ShowMessage(Format('%d Datensätze ausgegeben.', [I]));
      END;
  FINALLY
    Screen.Cursor := oldC;
  END;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ExportDatasetToFile(FileName : STRING;
  Dataset : TDataset;
  Delimiter : STRING;
  Append : Boolean;
  Spaltennamen : Boolean = True;
  TextInHochKomma : Boolean = True;
  DosZeichenSatz : Boolean = False);
VAR
  St : tFileStream;
  M : Word;
BEGIN
  //Append :=True;
  // 1. Datei erzeugen wenn sie nicht existiert
  IF ((fileexists(FileName)) AND
    (NOT Append)) OR
    (NOT fileexists(FileName)) THEN
    FileClose(FileCreate(FileName));

  IF Append THEN
    M := fmOpenReadWrite
  ELSE
    M := fmOpenWrite;
  M := M OR fmShareDenyWrite;

  St := tFileStream.Create(FileName, M);
  IF Append THEN
    ST.Position := ST.Size;

  TRY
    ExportDatasetToStream(St,
      Dataset,
      Delimiter,
      Spaltennamen,
      TextInHochKomma,
      DosZeichenSatz);
  FINALLY
    St.Free;
  END;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ExportDatasetToExcel(FileName : STRING; Dataset : TDataset; Spaltennamen : Boolean = True);

VAR
  XLSfile : TXLSfile;
  I, J : Integer;
  //    S         : string;
  FieldDesc : tField;
  OldC : tCursor;
BEGIN
  IF Dataset.RecordCount >= 65535 THEN
    ShowMessage(Format('ACHTUNG Excel kann nur 65535 Datensätze verarbeiten. '^M
      +
      'Ihre Selektion beträgt aber %d Datensätze.'^M +
      'Das sind %d die nicht verarbeitet werden.', [Dataset.RecordCount,
      Dataset.RecordCount - $FFFF]));

  OldC := Screen.Cursor;
  Screen.Cursor := crSqlWait;

  TRY
    XLSFile := TXLSfile.Create(Self);
    (*
      try
        scExcelExport1.LoadDefaultProperties;
        scExcelExport1.ExcelVisible:=False;
        scExcelExport1.WorksheetName := 'GERA Export';
        scExcelExport1.Dataset:=Dataset;
        scExcelExport1.ExportDataset;
        scExcelExport1.SaveAs(FileName,ffXLS);
        //scExcelExport1.SaveAs('c:\ExcelExport.htm',ffHTM);
        //scExcelExport1.SaveAs('c:\ExcelExport.csv',ffCSV);
      finally
        scExcelExport1.Disconnect;
      end;
    *)

    I := 1;
    WITH Dataset DO
      BEGIN
        // Feldnamen als Überschrift ausgeben !!
        IF Spaltennamen THEN
          FOR J := 0 TO FieldCount - 1 DO
            XLSFile.AddStrCell(J + 1, I, [], Fields[J].FieldName);

        First;
        WHILE (NOT EOF) AND (I < 65534) {Excel kann nur 65535 Zeilen !!!} DO
          BEGIN

            inc(i);
            //   If i > 4 then break;
            FOR J := 0 TO Fields.Count - 1 DO
              BEGIN
                FieldDesc := Fields[J];
                //  Nimmt Bildnamen aus BilderDatei 25.11.08
                IF Fields[J].FieldName = 'SHOP_IMAGE_LARGE' THEN
                  BEGIN
                    ArtikelBilderQuery.Sql.Clear;
                    ArtikelBilderQuery.Sql.Add('SELECT * FROM ARTIKEL_BILDER WHERE Artikel_ID=' + IntToStr(FieldByName('REC_ID').AsInteger));
                    ArtikelBilderQuery.Open;
                    XLSFile.AddStrCell(J + 1, I, [], ArtikelBilderQuery.FieldByName('BILD').AsString);
                  END
                ELSE
                  // 25.11.08
                  //XLSWrite.WriteString(Col - 1,Row - 1,-1,Grid.Cells[Col,Row])
      //              If (J =6 ) and (I < 4 ) then showmessage('datatype : '+inttostr(ord(FieldDesc.DataType ))+'  '+FieldDesc.AsString );
                  CASE FieldDesc.DataType OF
                    ftInteger,
                      ftSmallInt,
                      ftAutoInc,
                      ftLargeInt,
                      ftBCD : XLSFile.AddDoubleCell(J + 1, I, [], FieldDesc.AsInteger);
                    ftFloat,
                      ftCurrency : XLSFile.AddDoubleCell(J + 1, I, [], FieldDesc.AsFloat);
                    ftMemo : XLSFile.AddStrCell(J + 1, I, [], FieldDesc.AsString);
                    ELSE
                      XLSFile.AddStrCell(J + 1, I, [], FieldDesc.AsString);

                  END; //case
              END; //for
            Next;
          END; //while
      END; //with

    XLSFile.FileName := FileName;
    XLSFile.Write;

    IF (Dataset.RecordCount >= 65535) THEN
      BEGIN
        IF Dataset.RecordCount = 1 THEN // etwas rechtschreibung!
          ShowMessage(Format('%d Datensatz nach ' + FileName + ' ausgegeben.',
            [$FFFF]))
        ELSE
          ShowMessage(Format('%d Datensätze nach ' + FileName + ' ausgegeben.',
            [$FFFF]));
      END
    ELSE
      BEGIN
        IF Dataset.RecordCount = 1 THEN // etwas rechtschreibung!
          ShowMessage(Format('%d Datensatz nach ' + FileName + ' ausgegeben.',
            [Dataset.RecordCount]))
        ELSE
          ShowMessage(Format('%d Datensätze nach ' + FileName + ' ausgegeben.',
            [Dataset.RecordCount]));
      END;

  FINALLY
    //XLSFile.Free;

    Screen.Cursor := oldC;

  END;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.UpdateArtikelPreis(RechTyp, Artikel_ID, Addr_ID : Integer; Preis : Double);
BEGIN
  (*
                   UpdateArtikelPreis(EK_RECH,
                                      JPosTabARTIKEL_ID.Value,
                                      JourTabAddr_ID.Value,
                                      NewEK);
  *)
  Uniquery.Close;
  UniQuery.RequestLive := True;
  TRY
    UniQuery.Sql.Text := 'select * from ARTIKEL_PREIS where ARTIKEL_ID=' +
      IntToStr(Artikel_ID) +
      ' and ADRESS_ID=' +
      IntToStr(Addr_ID);
    UniQuery.Open;
    IF UniQuery.RecordCount = 1 THEN
      BEGIN
        UniQuery.Edit;
        TRY
          UniQuery.FieldByName('PREIS').AsFloat := Preis;
          UniQuery.FieldByName('GEAEND').AsDateTime := Now;
          UniQuery.FieldByName('GEAEND_NAME').AsString := View_User;

          UniQuery.Post;
        EXCEPT
          UniQuery.Cancel;
        END;

      END
    ELSE
      IF UniQuery.RecordCount = 0 THEN
        BEGIN
          // neuen Eintrag anlegen
          UniQuery.Append;
          TRY
            UniQuery.FieldByName('ARTIKEL_ID').AsInteger := Artikel_ID;
            UniQuery.FieldByName('ADRESS_ID').AsInteger := Addr_ID;
            UniQuery.FieldByName('PREIS_TYP').AsInteger := RechTyp;
            UniQuery.FieldByName('PREIS').AsFloat := Preis;
            UniQuery.FieldByName('GEAEND').AsDateTime := Now;
            UniQuery.FieldByName('GEAEND_NAME').AsString := View_User;

            UniQuery.Post;
          EXCEPT
            UniQuery.Cancel;
          END;
        END;
  FINALLY
    Uniquery.Close;
    UniQuery.RequestLive := False;
  END;
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.GetSearchSQL(Felder : ARRAY OF STRING; Suchbegriff : STRING) :
  STRING;
VAR
  Stk : TaaStringStack;
  Parser : TaaSearchParser;
  RPNNode : TaaRPNNode;
  Word1 : STRING;
  Word2 : STRING;
  SQL, SQL2, S : STRING;
  I, P : Integer;

BEGIN
  Stk := NIL;
  Parser := NIL;
  TRY
    Stk := TaaStringStack.Create;
    Parser := TaaSearchParser.Create(Suchbegriff);
    RPNNode := Parser.RPN;
    WHILE (RPNNode <> NIL) DO
      BEGIN
        IF (RPNNode IS TaaRPNWord) THEN
          Stk.Push('(@@@FN@@@ like ''%' + TaaRPNWord(RPNNode).PhraseWord +
            '%'')')
        ELSE
          IF (RPNNode IS TaaRPN_AND) THEN
            BEGIN
              Word2 := Stk.Pop;
              Word1 := Stk.Pop;
              Stk.Push('(' + Word1 + ' and ' + Word2 + ')');
            END
          ELSE
            IF (RPNNode IS TaaRPN_OR) THEN
              BEGIN
                Word2 := Stk.Pop;
                Word1 := Stk.Pop;
                Stk.Push('(' + Word1 + ' or ' + Word2 + ')');
              END
            ELSE
              BEGIN
                Word1 := Stk.Pop;
                Stk.Push('(not ' + Word1 + ')');
              END;
        RPNNode := RPNNode.Next;
      END;

    SQL := '';
    SQL2 := Stk.Pop; // in SQL2 steht jetzt der SQL-Befehl mit dem Feldplatzhalter

    // Feldnamen einfügen
    IF Length(Felder) > 0 THEN // Alle Felder durchlaufen
      BEGIN
        FOR i := 0 TO length(Felder) - 1 DO
          BEGIN
            S := SQL2;
            WHILE Pos('@@@FN@@@', S) > 0 DO // Platzhalter gefunden
              BEGIN
                P := Pos('@@@FN@@@', S);
                Delete(S, P, 8); // Platzhalter löschen
                insert(Felder[i], S, P); // und durch Feldnamen ersetzen
              END;

            // wenn in mehreren Feldern gesucht wird dan die einzelnen Suchen
            // mit oder verknüpfen
            IF length(SQL) > 0 THEN
              SQL := SQL + ' or ';
            SQL := SQL + '(' + S + ')';
          END;
      END;
    Result := SQL; // SQL-String zurückgeben
  FINALLY
    Stk.Free;
    Parser.Free;
  END;
END;
//----------------------Suchen nur ab Anfang -----------------------------------

FUNCTION TDM1.GetSearchSQL2(Felder : ARRAY OF STRING; Suchbegriff : STRING) :
  STRING;
VAR
  Stk : TaaStringStack;
  Parser : TaaSearchParser;
  RPNNode : TaaRPNNode;
  Word1 : STRING;
  Word2 : STRING;
  SQL, SQL2, S : STRING;
  I, P : Integer;

BEGIN
  Stk := NIL;
  Parser := NIL;
  TRY
    Stk := TaaStringStack.Create;
    Parser := TaaSearchParser.Create(Suchbegriff);
    RPNNode := Parser.RPN;
    WHILE (RPNNode <> NIL) DO
      BEGIN
        IF (RPNNode IS TaaRPNWord) THEN
          Stk.Push('(FN@@@ like ''%' + TaaRPNWord(RPNNode).PhraseWord + '%'')')
        ELSE
          IF (RPNNode IS TaaRPN_AND) THEN
            BEGIN
              Word2 := Stk.Pop;
              Word1 := Stk.Pop;
              Stk.Push('(' + Word1 + ' and ' + Word2 + ')');
            END
          ELSE
            IF (RPNNode IS TaaRPN_OR) THEN
              BEGIN
                Word2 := Stk.Pop;
                Word1 := Stk.Pop;
                Stk.Push('(' + Word1 + ' or ' + Word2 + ')');
              END
            ELSE
              BEGIN
                Word1 := Stk.Pop;
                Stk.Push('(not ' + Word1 + ')');
              END;
        RPNNode := RPNNode.Next;
      END;

    SQL := '';
    SQL2 := Stk.Pop; // in SQL2 steht jetzt der SQL-Befehl mit dem Feldplatzhalter

    // Feldnamen einfügen
    IF Length(Felder) > 0 THEN // Alle Felder durchlaufen
      BEGIN
        FOR i := 0 TO length(Felder) - 1 DO
          BEGIN
            S := SQL2;
            WHILE Pos('@@@FN@@@', S) > 0 DO // Platzhalter gefunden
              BEGIN
                P := Pos('@@@FN@@@', S);
                Delete(S, P, 8); // Platzhalter löschen
                insert(Felder[i], S, P); // und durch Feldnamen ersetzen
              END;

            // wenn in mehreren Feldern gesucht wird dan die einzelnen Suchen
            // mit oder verknüpfen
            IF length(SQL) > 0 THEN
              SQL := SQL + ' or ';
            SQL := SQL + '(' + S + ')';
          END;
      END;
    Result := SQL; // SQL-String zurückgeben
  FINALLY
    Stk.Free;
    Parser.Free;
  END;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.FirmaTabCalcFields(DataSet : TDataSet);
BEGIN
  FirmaTabUSER_AKT.AsString := View_User;
  FirmaTabLEITWAEHRUNG.AsString := Leitwaehrung;
  FirmaTabMANDANT_NAME.AsString := AktMandant;
END;

//------------------------------------------------------------------------------

FUNCTION TDM1.GetVertreterProv(ID : Integer) : Double;
BEGIN
  Result := 0;
  IF ID < 1 THEN
    exit;

  IF ID = LastVertrID THEN
    BEGIN
      Result := LastVertrProz;
      exit;
    END;

  IF NOT VertreterTab.Active THEN
    VertreterTab.Open;

  IF (VertreterTab.Locate('VERTRETER_ID', ID, [])) AND
    (VertreterTab.FieldByName('VERTRETER_ID').AsInteger = ID) THEN
    BEGIN
      LastVertrID := VertreterTab.FieldByName('VERTRETER_ID').AsInteger;
      LastVertrProz := VertreterTab.FieldByName('PROVISIONSATZ').AsFloat;

      Result := LastVertrProz;
    END;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.CheckTerminalID;
VAR
  I : Integer;
BEGIN
  Uniquery.SQL.Text := 'select SUBSTRING_INDEX(USER(),"@",-1) as PC';
  Uniquery.Open;
  TermIDStr := Uniquery.FieldByName('PC').AsString;
  Uniquery.Close;

  Uniquery.SQL.Text := 'select MAINKEY, NAME, VAL_INT from REGISTRY ' +
    'where MAINKEY="MAIN\\TERMINAL" and NAME="' + TermIDStr + '"';
  Uniquery.Open;
  IF Uniquery.RecordCount = 1 THEN
    BEGIN
      TermID := Uniquery.FieldByName('VAL_INT').AsInteger;
      Uniquery.Close;
    END
  ELSE
    BEGIN
      Uniquery.Close;
      Uniquery.SQL.Text := 'select COUNT(*) as ANZ FROM REGISTRY ' +
        'where MAINKEY="MAIN\\TERMINAL"';
      UniQuery.Open;
      I := UniQuery.FieldByName('ANZ').AsInteger + 1;
      UniQuery.Close;

      UniQuery.SQL.Text := 'INSERT INTO REGISTRY SET ' +
        'MAINKEY="MAIN\\TERMINAL", NAME="' + TermIDStr + '", ' +
        'VAL_INT=' + IntToStr(i) + ', VAL_TYP=3';
      UniQuery.ExecSQL;
      TermID := I;
    END;
  //CheckRegister2;
END;
//------------------------------------------------------------------------------

FUNCTION tDM1.GetArtikelPreis(ArtikelID, KunID, PE : Integer;
  Brutto : Boolean;
  Menge : Double;
  VAR Preis : Double) : Boolean;

VAR
  PR, M, B : Double;
  PreisTyp : Integer;
BEGIN
  Result := False;

  IF PE = 0 THEN
    PreisTyp := EK_RECH
  ELSE
    PreisTyp := VK_RECH;

  IF (ArtInfoTab.Active) AND
    (ArtInfoTab.RecordCount > 0) AND
    ((DM1.ArtInfoTab.Locate('REC_ID;PREIS_TYP;ADRESS_ID', VarArrayOF([ArtikelID, PreisTyp, KunID]), [])) OR
    (DM1.ArtInfoTab.Locate('REC_ID', ArtikelID, []))) THEN
    //(ArtInfoTab.Locate ('REC_ID',ArtikelID,[])) then
    BEGIN
      IF Brutto THEN
        BEGIN
          //Kundenpreis prüfen
          IF (ArtInfoTabAdress_ID.AsInteger = KunID) AND
            (ArtInfoTabPreis.AsFloat <> 0) THEN
            BEGIN
              PR := ArtInfoTabPreis.AsFloat;
              // Brutto berechnen
              CASE ArtInfoTabSTEUER_CODE.AsInteger OF
                1 : M := MWSTTab[1];
                2 : M := MWSTTab[2];
                3 : M := MWSTTab[3];
                ELSE
                  M := MWSTTab[0];
              END;
              M := M + 100; // jetzt z.B. 116

              B := CAO_round(PR * M); // jetzt ganze Cent
              B := CAO_Round(B / BR_RUND_WERT) * BR_RUND_WERT / 100;
              Preis := B;
            END
          ELSE
            BEGIN
              CASE PE OF
                0 : PR := ArtInfoTabEK_PREIS.Value;
                1 : PR := ArtInfoTabVK1B.Value;
                2 : PR := ArtInfoTabVK2B.Value;
                3 : PR := ArtInfoTabVK3B.Value;
                4 : PR := ArtInfoTabVK4B.Value;
                ELSE
                  PR := ArtInfoTabVK5B.Value;
              END;
              Preis := PR;
            END;

          IF (ArtInfoTabRABGRP_ID.AsString <> '-') AND
            (length(ArtInfoTabRABGRP_ID.AsString) > 0) AND
            //(ArtInfoTabPreis.AsFloat=0) then
          (ArtInfoTabAdress_ID.AsInteger <> KunID) THEN
            BEGIN
              // Ist ein Artikel mit Rabattgruppe !!!
              // Brutto muß berechnet werden

              // Listenpreis festlegen
              CASE AnzPreis OF
                1 : PR := ArtInfoTabVK1.Value;
                2 : PR := ArtInfoTabVK2.Value;
                3 : PR := ArtInfoTabVK3.Value;
                4 : PR := ArtInfoTabVK4.Value;
                ELSE
                  PR := ArtInfoTabVK5.Value;
              END;

              // nur wenn Rab.Gruppe gefunden
              IF CalcRabGrpPreis(ArtInfoTabRABGRP_ID.AsString, PE, PR) THEN
                BEGIN
                  CASE ArtInfoTabSTEUER_CODE.AsInteger OF
                    1 : M := MWSTTab[1];
                    2 : M := MWSTTab[2];
                    3 : M := MWSTTab[3];
                    ELSE
                      M := MWSTTab[0];
                  END;
                  M := M + 100; // jetzt z.B. 116

                  B := CAO_round(PR * M); // jetzt ganze Cent
                  B := CAO_Round(B / BR_RUND_WERT) * BR_RUND_WERT / 100;
                  Preis := B;
                END;
            END;
        END
      ELSE
        BEGIN
          // Netto
          //Kundenpreis prüfen
          IF (ArtInfoTabAdress_ID.AsInteger = KunID) AND
            (ArtInfoTabPreis.AsFloat <> 0) THEN
            BEGIN
              // Kundenpreis verwenden
              Preis := ArtInfoTabPreis.AsFloat;
            END
          ELSE
            BEGIN
              // normalen Preis verwenden
              CASE PE OF
                0 : PR := ArtInfoTabEK_Preis.Value;
                1 : PR := ArtInfoTabVK1.Value;
                2 : PR := ArtInfoTabVK2.Value;
                3 : PR := ArtInfoTabVK3.Value;
                4 : PR := ArtInfoTabVK4.Value;
                ELSE
                  PR := ArtInfoTabVK5.Value;
              END;
              Preis := PR;
            END;

          IF (ArtInfoTabRABGRP_ID.AsString <> '-') AND
            (length(ArtInfoTabRABGRP_ID.AsString) > 0) AND
            (ArtInfoTabAdress_ID.AsInteger <> KunID) THEN
            BEGIN
              // Ist ein Artikel mit Rabattgruppe !!!

              // Listenpreis festlegen
              CASE AnzPreis OF
                1 : PR := ArtInfoTabVK1.Value;
                2 : PR := ArtInfoTabVK2.Value;
                3 : PR := ArtInfoTabVK3.Value;
                4 : PR := ArtInfoTabVK4.Value;
                ELSE
                  PR := ArtInfoTabVK5.Value;
              END;

              // nur wenn Rab.Gruppe gefunden
              IF CalcRabGrpPreis(ArtInfoTabRABGRP_ID.AsString, PE, PR) THEN
                Preis := PR;
            END;

          // -----------------------------------
          // Staffelpreis ermitteln NH20050311-1
          // -----------------------------------
          //
          IF (Menge >= ArtInfoTabMENGE5.Value) AND
            (ArtInfoTabPREIS5.value > 0) THEN
            BEGIN
              CASE AnzPreis OF
                1 : PR := (ArtInfoTabVK1.Value * 100 - ArtInfoTabVK1.Value * ArtInfoTabPREIS5.value) / 100;
                2 : PR := (ArtInfoTabVK2.Value * 100 - ArtInfoTabVK2.Value * ArtInfoTabPREIS5.value) / 100;
                3 : PR := (ArtInfoTabVK3.Value * 100 - ArtInfoTabVK3.Value * ArtInfoTabPREIS5.value) / 100;
                4 : PR := (ArtInfoTabVK4.Value * 100 - ArtInfoTabVK4.Value * ArtInfoTabPREIS5.value) / 100;
                ELSE
                  PR := (ArtInfoTabVK5.Value * 100 - ArtInfoTabVK5.Value * ArtInfoTabPREIS5.value) / 100;
              END;
              Preis := CAO_round_nk(PR, DM1.VK_NACHKOMMA);
            END
          ELSE
            IF (Menge >= ArtInfoTabMENGE4.Value) AND
              (ArtInfoTabPREIS4.value > 0) THEN
              BEGIN
                CASE AnzPreis OF
                  1 : PR := (ArtInfoTabVK1.Value * 100 - ArtInfoTabVK1.Value * ArtInfoTabPREIS4.value) / 100;
                  2 : PR := (ArtInfoTabVK2.Value * 100 - ArtInfoTabVK2.Value * ArtInfoTabPREIS4.value) / 100;
                  3 : PR := (ArtInfoTabVK3.Value * 100 - ArtInfoTabVK3.Value * ArtInfoTabPREIS4.value) / 100;
                  4 : PR := (ArtInfoTabVK4.Value * 100 - ArtInfoTabVK4.Value * ArtInfoTabPREIS4.value) / 100;
                  ELSE
                    PR := (ArtInfoTabVK5.Value * 100 - ArtInfoTabVK5.Value * ArtInfoTabPREIS4.value) / 100;
                END;
                Preis := CAO_round_nk(PR, DM1.VK_NACHKOMMA);
              END
            ELSE
              IF (Menge >= ArtInfoTabMENGE3.Value) AND
                (ArtInfoTabPREIS3.value > 0) THEN
                BEGIN
                  CASE AnzPreis OF
                    1 : PR := (ArtInfoTabVK1.Value * 100 - ArtInfoTabVK1.Value * ArtInfoTabPREIS3.value) / 100;
                    2 : PR := (ArtInfoTabVK2.Value * 100 - ArtInfoTabVK2.Value * ArtInfoTabPREIS3.value) / 100;
                    3 : PR := (ArtInfoTabVK3.Value * 100 - ArtInfoTabVK3.Value * ArtInfoTabPREIS3.value) / 100;
                    4 : PR := (ArtInfoTabVK4.Value * 100 - ArtInfoTabVK4.Value * ArtInfoTabPREIS3.value) / 100;
                    ELSE
                      PR := (ArtInfoTabVK5.Value * 100 - ArtInfoTabVK5.Value * ArtInfoTabPREIS3.value) / 100;
                  END;
                  Preis := CAO_round_nk(PR, DM1.VK_NACHKOMMA);
                END
              ELSE
                IF (Menge >= ArtInfoTabMENGE2.Value) AND
                  (ArtInfoTabPREIS2.value > 0) THEN
                  BEGIN
                    CASE AnzPreis OF
                      1 : PR := (ArtInfoTabVK1.Value * 100 - ArtInfoTabVK1.Value * ArtInfoTabPREIS2.value) / 100;
                      2 : PR := (ArtInfoTabVK2.Value * 100 - ArtInfoTabVK2.Value * ArtInfoTabPREIS2.value) / 100;
                      3 : PR := (ArtInfoTabVK3.Value * 100 - ArtInfoTabVK3.Value * ArtInfoTabPREIS2.value) / 100;
                      4 : PR := (ArtInfoTabVK4.Value * 100 - ArtInfoTabVK4.Value * ArtInfoTabPREIS2.value) / 100;
                      ELSE
                        PR := (ArtInfoTabVK5.Value * 100 - ArtInfoTabVK5.Value * ArtInfoTabPREIS2.value) / 100;
                    END;
                    Preis := CAO_round_nk(PR, DM1.VK_NACHKOMMA);
                  END;

        END;

      Result := True;
    END;
END;
//------------------------------------------------------------------------------
// Liefert True zurück wenn in der Warengruppe ein Kalkulationsfaktor festgelegt
// wurde
//------------------------------------------------------------------------------

FUNCTION tDM1.GetWGRCalcFaktor(Wgr, PreisID : Integer; VAR Faktor : Double) :
  Boolean;
VAR
  I : Integer;
BEGIN
  Result := False;
  IF (PreisID < 1) OR (PreisID > 5) THEN
    exit; // VK geht nur von 1-5 !!!

  IF WgrFaktorCache.Wgr = Wgr THEN
    BEGIN
      Faktor := WgrFaktorCache.FTab[PreisID];
    END
  ELSE
    BEGIN
      IF NOT WgrTab.Active THEN
        WgrTab.Open;

      IF (WgrTab.RecordCount > 0) AND
        (WgrTab.Locate('ID', Wgr, [])) THEN
        BEGIN
          FOR i := 1 TO AnzPreis DO
            BEGIN
              WgrFaktorCache.FTab[i] := WgrTab.FieldByName('VK' + IntToStr(I) +
                '_FAKTOR').AsFloat;
            END;
          WgrFaktorCache.Wgr := Wgr;
          Faktor := WgrFaktorCache.FTab[PreisID];
        END
      ELSE
        WgrFaktorCache.Wgr := -1; // Cache ungültig da Wgr nicht gefunden

    END;
  Result := Faktor <> 0;
END;
//------------------------------------------------------------------------------

FUNCTION GetProjectVersion : STRING;
VAR
  Null, InfoSize, FixInfo : DWord;
  PFixInfo : PVSFixedFileInfo;
  Zeiger : Pointer;
BEGIN
  result := '??';
  Null := 0;
  InfoSize := GetFileVersionInfoSize(PChar(application.exename), Null);
  IF InfoSize > 0 THEN
    BEGIN
      Zeiger := GetMemory(InfoSize);
      TRY
        IF assigned(Zeiger) THEN
          BEGIN
            GetFileVersionInfo(PChar(application.exename), Null, InfoSize,
              Zeiger);
            IF VerQueryValue(Zeiger, '\', Pointer(PFixInfo), FixInfo) THEN
              BEGIN
                result := Format('%d.%d.%d.%d',
                  [HiWord(PFixInfo^.dwFileVersionMS),
                  LoWord(PFixInfo^.dwFileVersionMS),
                    HiWord(PFixInfo^.dwFileVersionLS),
                    LoWord(PFixInfo^.dwFileVersionLS)]);
              END;
          END;
      FINALLY
        FreeMemory(Zeiger);
      END;
    END;
END;
//------------------------------------------------------------------------------
// Benutzerrechte
//------------------------------------------------------------------------------

FUNCTION TDM1.CaoSecurityFindUser(UserName : STRING; VAR MA_ID,
  GRUPPE_ID : Integer) : Boolean;
BEGIN
  // Vorgehen :
  // 1. Mitarbeiter über seinen Namen in der Tabelle MITARBEITER suchen
  // 2. wenn gefunden, die MITARBEITER.MA_ID verwenden umaus der Tabelle
  //    BENUTZERRECHTE die zugehörige Gruppe zu ermitteln
  // 3. MA_ID und Gruppe zurückliefern,
  //    wenn User nicht gefunden, Result=False

//     Result    :=False;
  GRUPPE_ID := -1;
  MA_ID := -1;

  UniQuery.Close;
  UniQuery.SQL.Text :=
    'SELECT MA_ID, Anrede, Anzeige_Name,GruppeID, ObjGruppeID, LOGIN_NAME ,Kurz, Zweig FROM TBL_KONTAKTE ' +
    'WHERE LOGIN_NAME=:NAME';
  UniQuery.ParamByName('NAME').AsString := UserName;
  UniQuery.Open;

  IF UniQuery.RecordCount = 1 THEN
    BEGIN
      MA_ID := UniQuery.FieldByName('MA_ID').AsInteger;
      MitarbeiterID := UniQuery.FieldByName('MA_ID').AsInteger;
      MitarbeiterName := UniQuery.FieldByName('Anzeige_name').AsString;
      MitarbeiterKurz := UniQuery.FieldByName('Kurz').AsString;
      MitarbeiterAnrede := UniQuery.FieldByName('Anrede').AsString;
      MitarbeiterGruppeID := UniQuery.FieldByName('GruppeID').AsInteger;
      MitarbeiterObjGruppeID := UniQuery.FieldByName('ObjGruppeID').AsInteger;
      Zweig := UniQuery.FieldByName('Zweig').AsString;
      UniQuery.Close;
      UniQuery.SQL.Text :='SELECT stufe,eintritt FROM tbl_lizenzstufe WHERE mitarbeiter="'+inttostr(Mitarbeiterid)+'" ';
      UniQuery.Open;
      IF UniQuery.RecordCount = 1 THEN
        MitarbeiterLizenzstufe := UniQuery.FieldByName('stufe').AsInteger
      else
        showmessage('Keine gültige Lizenzstufe gefunden !');
      UniQuery.Close;


{$IFDEF ALPHA}
      //Jetzt Benutzergruppe finden
      UniQuery.Sql.Text :=
        'SELECT GRUPPEN_ID, USER_ID FROM BENUTZERRECHTE ' +
        'WHERE USER_ID=' + IntToStr(MA_ID) + ' AND ' +
        'MODUL_ID=0 and SUBMODUL_ID=0';
      UniQuery.Open;
      IF UniQuery.RecordCount = 1 THEN
        BEGIN
          GRUPPE_ID := UniQuery.FieldByName('GRUPPEN_ID').AsInteger;

          Result := True;
        END;
      UniQuery.Close;

{$ELSE}
      GRUPPE_ID := 1;
      Result := True;
{$ENDIF}
    END
  ELSE
    BEGIN
      UniQuery.Close;
      Result := False;
    END;
  IF MitarbeiterTab.Lookup('idmitarbeiter', Mitarbeiterid, 'objektkontingent') > 0 THEN
    BEGIN
      Uniquery.Sql.Text := 'select objnr from tbl_obj where idmitarbeiter ="' + IntToStr(Mitarbeiterid) + '" and Status = 0';
      Uniquery.Open;
      IF Uniquery.RecordCount > MitarbeiterTab.Lookup('idmitarbeiter', Mitarbeiterid, 'objektkontingent') THEN
        showmessage('Ihre aktiven Objekte überschreiten Ihr Kontingent !');
      Uniquery.close;
    END;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.CaoSecurityLoadGruppeRechte(Sender : TObject);
BEGIN
{$IFDEF ALPHA}
  WITH UniQuery DO
    BEGIN
      Close;
      Sql.Text :=
        'select GRUPPEN_ID, MODUL_ID, SUBMODUL_ID, RECHTE ' +
        'from BENUTZERRECHTE B ' +
        'where GRUPPEN_ID=' + IntToStr(CaoSecurity.CurrGroupID) + ' and ' +
        'USER_ID=-1 and MODUL_ID > 0 ' +
        'order by GRUPPEN_ID, MODUL_ID, SUBMODUL_ID';
      Open;
      WHILE NOT Eof DO
        BEGIN
          CaoSecurity.AddGruppenRecht(FieldByName('MODUL_ID').AsInteger,
            FieldByName('SUBMODUL_ID').AsInteger,
            FieldByName('RECHTE').AsInteger);

          Next;
        END;
      Close;
    END;
{$ENDIF}
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.CaoSecurityLoadUserRechte(Sender : TObject);
BEGIN
{$IFDEF ALPHA}
  WITH UniQuery DO
    BEGIN
      Close;
      Sql.Text :=
        'select GRUPPEN_ID,USER_ID,MODUL_ID,SUBMODUL_ID,RECHTE ' +
        'from BENUTZERRECHTE where ' +
        'USER_ID=' + IntToStr(CaoSecurity.CurrUserID) +
        ' and GRUPPEN_ID=-1 and MODUL_ID > 0 ' +
        'order by MODUL_ID, SUBMODUL_ID';
      Open;
      WHILE NOT Eof DO
        BEGIN
          CaoSecurity.AddUserRecht(FieldByName('MODUL_ID').AsInteger,
            FieldByName('SUBMODUL_ID').AsInteger,
            FieldByName('RECHTE').AsInteger);

          Next;
        END;
      Close;
    END;
{$ENDIF}
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.StarteNewProgramm(Name, Cmd, Dir : STRING);
VAR
  res : integer;
  Msg : STRING;
BEGIN
  res := shellexecute(application.mainform.handle,
    pchar('open'), PChar(name), pchar(cmd), pchar(dir), sw_shownormal);

  IF res <= 32 THEN
    BEGIN
      Msg := GetErrorStr(Res);
      messagebox(application.mainform.handle, @msg[1], 'Problem', 16);
    END;
END;
//------------------------------------------------------------------------------

//------------------------------------------------------------------------------
// Funktionen um ein virtuelles Recordlocking zu ermöglichen
//------------------------------------------------------------------------------
{
// Problem, Funktion wird erst ab MySQL 4.0.2 unterstützt :-(

function TDM1.CaoSecurityIsFreeLock(ModulID: String;
  SatzID: Integer): Boolean;
begin
     LockQuery.Close;
     LockQuery.Sql.Text :='SELECT IS_FREE_LOCK("'+ModulID+'_RECID_'+
                          IntToStr(SatzID)+'") as LOCKVAR';
     LockQuery.Open;
     if (LockQuery.RecordCount=1)and
        (LockQuery.FieldByName('LOCKVAR').AsInteger=1)
      then Result :=True
      else Result :=False;

     LockQuery.Close;
end;
}
//------------------------------------------------------------------------------

FUNCTION TDM1.CaoSecuritySetLock(ModulID : STRING;
  SatzID : Integer) : Boolean;
BEGIN
  LockQuery.Close;
  LockQuery.Sql.Text := 'SELECT GET_LOCK("' + ModulID + '_RECID_' +
    IntToStr(SatzID) + '",3) as LOCKVAR';
  LockQuery.Open;
  IF (LockQuery.RecordCount = 1) AND
    (LockQuery.FieldByName('LOCKVAR').AsInteger = 1) THEN
    Result := True
  ELSE
    Result := False;

  LockQuery.Close;
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.CaoSecurityReleaseLock(ModulID : STRING;
  SatzID : Integer) : Boolean;
BEGIN
  LockQuery.Close;
  LockQuery.Sql.Text := 'SELECT RELEASE_LOCK("' + ModulID + '_RECID_' +
    IntToStr(SatzID) + '") as LOCKVAR';
  LockQuery.Open;
  IF (LockQuery.RecordCount = 1) AND
    (LockQuery.FieldByName('LOCKVAR').AsInteger = 1) THEN
    Result := True
  ELSE
    Result := False;

  LockQuery.Close;
END;
//------------------------------------------------------------------------------

PROCEDURE TDM1.LockError(Error : Integer);
BEGIN
  //MessageDlg (_('Datensatz-Sperre Code : ')+IntToStr(Error),mterror,[mbok],0);
  MessageDlg(_('der Datensatz wird zur Zeit von ' +
    'einem anderen Nutzer bearbeitet,' + #13#10 +
    'warten Sie bis der Benutzer den Datensatz ' +
    'freigegebn hat.' + #13#10 +
    'Code : ') + IntToStr(Error),
    mterror, [mbok], 0);
END;
//------------------------------------------------------------------------------

FUNCTION TDM1.InsertStuecklistenArtikel(JournalID,
  JournalposID,
  ArtikelID,
  AddrID,
  BelegArt : Integer;
  Menge : Double;
  BelegNum : STRING) : Boolean;

VAR
  IStr : STRING;
BEGIN
  Result := False;

  STListTab.Close;
  STListTab.ParamByName('ID').ASInteger := ArtikelID;
  STListTab.Open;

  WHILE NOT STListTab.Eof DO
    BEGIN
      // Menge erniedrigen
      {
      ArtMengeTab.Close;
      ArtMengeTab.ParamByName ('ID').Value :=STListTabART_ID.Value;
      ArtMengeTab.ParamByName ('SUBMENGE').Value :=JPosTabMenge.Value * STListTabMENGE.Value;
      ArtMengeTab.ExecSql;
      }
      IF length(IStr) > 0 THEN
        IStr := IStr + ';' + #13#10;

      // Batch-SQL erzeugen und die Stücklistenartikel mit in die Rechnung zu speichern
      // mit Artikeltyp="X"

      uniquery.close;
      uniquery.sql.text := 'select MATCHCODE,ARTNUM,BARCODE,LAENGE,' +
        'GROESSE,DIMENSION,GEWICHT,ME_EINHEIT,' +
        'LANGNAME, SN_FLAG ' +
        'from ARTIKEL where REC_ID=' +
        IntToStr(STListTabART_ID.Value);
      uniquery.open;

      IStr := Istr +
        'INSERT INTO JOURNALPOS SET ' +
        'QUELLE=' + IntToStr(BelegArt) +
        ',QUELLE_SUB=' + IntToStr(1) +
        ',JOURNAL_ID=' + IntToStr(JournalID) +
        ',ARTIKELTYP="X"' +
        ',ARTIKEL_ID=' + IntToStr(STListTabART_ID.Value) +
        ',TOP_POS_ID=' + IntToStr(JournalposID) +
        ',ADDR_ID=' + IntToStr(AddrID) +
        ',VRENUM="' + BelegNum + '"' +
        ',MENGE="' + FloatToStrEx(Menge * STListTabMENGE.Value) + '"' +
        ',POSITION=' + IntToStr(0) +
        ',MATCHCODE="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('MATCHCODE').AsString) + '"'
        +
        ',ARTNUM="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('ARTNUM').AsString) + '"' +
        ',BARCODE="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('BARCODE').AsString) + '"' +
        ',LAENGE="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('LAENGE').AsString) + '"' +
        ',GROESSE="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('GROESSE').AsString) + '"' +
        ',DIMENSION="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('DIMENSION').AsString) + '"'
        +
        ',GEWICHT=' + FloatToStrEx(uniquery.fieldbyname('GEWICHT').AsFloat) +
        ',ME_EINHEIT="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('ME_EINHEIT').AsString) + '"'
        +
        ',BEZEICHNUNG="' +
        ZSqlTypes.StringToSql(uniquery.fieldbyname('LANGNAME').AsString) + '"' +
        ',SN_FLAG="' + uniquery.fieldbyname('SN_FLAG').AsString + '"';

      uniquery.close;
      STListTab.Next;
    END;
  STListTab.Close;

  // Stücklistenunterartikel hinzufügen
  IF length(IStr) > 0 THEN
    BEGIN
      ZBatchSql1.Sql.Text := ISTr;
      TRY
        ZBatchSql1.ExecSql;
      EXCEPT
        MessageDlg(_('Fehler beim hinzufügen der Stücklisten-Unterartikel.'),
          mterror, [mbok], 0);
      END;
    END;
END;

//------------------------------------------------------------------------------

FUNCTION TDM1.UpdateStuecklistenArtikel(JournalID,
  JournalposID,
  ArtikelID : Integer;
  Menge : Double) : Boolean;
VAR
  IStr : STRING;
BEGIN
  Result := False;

  STListTab.Close;
  STListTab.ParamByName('ID').ASInteger := ArtikelID;
  STListTab.Open;

  WHILE NOT STListTab.Eof DO
    BEGIN
      IF length(IStr) > 0 THEN
        IStr := IStr + ';' + #13#10;

      // Batch-SQL erzeugen und die Stücklistenartikel (Menge) zu aktualisieren
      IStr := Istr +
        'UPDATE JOURNALPOS SET ' +
        'MENGE="' + FloatToStrEx(Menge * STListTabMENGE.Value) + '" ' +
        'WHERE JOURNAL_ID=' + IntToStr(JournalID) + ' ' +
        'AND TOP_POS_ID=' + IntToStr(JournalposID) + ' ' +
        'AND ARTIKEL_ID=' + IntToStr(STListTabART_ID.AsInteger);

      STListTab.Next;
    END;
  STListTab.Close;

  // Stücklistenunterartikel updaten
  IF length(IStr) > 0 THEN
    BEGIN
      DM1.ZBatchSql1.Sql.Text := ISTr;
      TRY
        ZBatchSql1.ExecSql;
      EXCEPT
        MessageDlg(_('Fehler beim aktualisieren der Stücklisten-Unterartikel.'),
          mterror, [mbok], 0);
      END;
    END;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.WgrTabNewRecord(DataSet : TDataSet);
BEGIN
  WgrTabSTEUER_CODE.Value := DefMwStCD;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ExportCSVDatasetToExcel(FileName : STRING;
  Dataset : TDataset;
  Spaltennamen : Boolean);
VAR
  XLSfile : TXLSfile;
  I, J : Integer;
  FieldDesc : tField;
  OldC : tCursor;
  Spalte : Integer;
  RecID : Integer;

BEGIN
  IF Dataset.RecordCount >= 65535 THEN
    ShowMessage(Format('ACHTUNG Excel vor 2007 kann nur 65535 Datensätze verarbeiten. '^M +
      'Ihre selerktion beträgt aber %d Datensätze.', [Dataset.RecordCount]));

  OldC := Screen.Cursor;
  Screen.Cursor := crSqlWait;

  XLSFile := TXLSfile.Create(Self);

  TRY
    I := 1;
    Spalte := 0;
    WITH Dataset DO
      BEGIN
        // Feldnamen als Überschrift ausgeben !!
        IF Spaltennamen THEN
          BEGIN
            FOR J := 0 TO FieldCount - 1 DO
              IF (Fields[J].FieldName = 'MATCHCODE') OR
                (Fields[J].FieldName = 'WARENGRUPPE') OR
                (Fields[J].FieldName = 'KURZNAME') OR
                (Fields[J].FieldName = 'LANGNAME') OR
                (Fields[J].FieldName = 'EK_PREIS') OR
                (Fields[J].FieldName = 'SHOP_IMAGE_LARGE') THEN
                BEGIN
                  XLSFile.AddStrCell(Spalte + 1, I, [], Fields[J].FieldName);
                  Inc(Spalte);
                END;
          END;

        First;
        WHILE (NOT EOF) AND (I < 65534) {Excel kann nur 65535 Zeilen !!!} DO
          BEGIN
            inc(i);
            Spalte := 0;
            FOR J := 0 TO Fields.Count - 1 DO
              IF (Fields[J].FieldName = 'MATCHCODE') OR
                (Fields[J].FieldName = 'WARENGRUPPE') OR
                (Fields[J].FieldName = 'KURZNAME') OR
                (Fields[J].FieldName = 'LANGNAME') OR
                (Fields[J].FieldName = 'EK_PREIS') OR
                (Fields[J].FieldName = 'SHOP_IMAGE_LARGE') THEN
                BEGIN
                  FieldDesc := Fields[J];

                  CASE FieldDesc.DataType OF
                    ftInteger,
                      ftSmallInt,
                      ftAutoInc,
                      ftLargeInt,
                      ftBCD : XLSFile.AddDoubleCell(Spalte + 1, I, [],
                        FieldDesc.AsInteger);
                    ftFloat,
                      ftCurrency : XLSFile.AddDoubleCell(Spalte + 1, I, [],
                        FieldDesc.AsFloat);
                    ELSE
                      IF (Fields[J].FieldName = 'SHOP_IMAGE_LARGE') THEN
                        BEGIN
                          ArtikelBilderQuery.Close;
                          ArtikelBilderQuery.Sql.Clear;
                          RecID := Dataset.FieldByName('REC_ID').AsInteger;
                          ArtikelBilderQuery.Sql.Text :=
                            'SELECT * FROM ARTIKEL_BILDER WHERE ' +
                            'ARTIKEL_ID=' + IntToStr(RecID);
                          ArtikelBilderQuery.Open;
                          IF (ArtikelBilderQuery.RecordCount > 0) THEN
                            XLSFile.AddStrCell(Spalte + 1, I, [],
                              ExtractFileName(ArtikelBilderQuery.FieldByName('BILD').AsString));
                        END
                      ELSE
                        XLSFile.AddStrCell(Spalte + 1, I, [],
                          FieldDesc.AsString);
                  END;
                  Inc(Spalte);
                END;
            Next;
          END;
      END;

    XLSFile.FileName := FileName;
    XLSFile.Write;
  FINALLY
    XLSFile.Free;
    Screen.Cursor := oldC;
  END;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ExportCSVDatasetToFile(FileName : STRING;
  Dataset : TDataset;
  Delimiter : STRING;
  Append,
  Spaltennamen,
  TextInHochKomma,
  DosZeichenSatz : Boolean);
VAR
  St : tFileStream;
  M : Word;
BEGIN
  //Append :=True;
  // 1. Datei erzeugen wenn sie nicht existiert
  IF ((fileexists(FileName)) AND (NOT Append)) OR
    (NOT fileexists(FileName)) THEN
    FileClose(FileCreate(FileName));

  IF Append THEN
    M := fmOpenReadWrite
  ELSE
    M := fmOpenWrite;

  M := M OR fmShareDenyWrite;

  St := tFileStream.Create(FileName, M);
  IF Append THEN
    ST.Position := ST.Size;
  TRY
    ExportCSVDatasetToStream(St,
      Dataset,
      Delimiter,
      Spaltennamen,
      TextInHochKomma,
      DosZeichenSatz);
  FINALLY
    St.Free;
  END;
END;

//------------------------------------------------------------------------------

PROCEDURE TDM1.ExportCSVDatasetToStream(Stream : TStream;
  Dataset : TDataset;
  Delimiter : STRING;
  Spaltennamen,
  TextInHochKomma,
  DosZeichenSatz : Boolean);
VAR
  I, J : Integer;
  Buffer, S : STRING;
  FieldDesc : tField;
  OldC : tCursor;
  RecID : Integer;
BEGIN
  OldC := Screen.Cursor;
  Screen.Cursor := crSqlWait;
  TRY
    WITH Dataset DO
      BEGIN
        // Feldnamen als Überschrift ausgeben !!
        IF (Fields.Count > 0) AND (Stream.Position = 0) AND (Spaltennamen) THEN
          BEGIN
            Buffer := '';
            FOR J := 0 TO FieldCount - 1 DO
              IF (Fields[J].FieldName = 'MATCHCODE') OR
                (Fields[J].FieldName = 'WARENGRUPPE') OR
                (Fields[J].FieldName = 'KURZNAME') OR
                (Fields[J].FieldName = 'LANGNAME') OR
                (Fields[J].FieldName = 'EK_PREIS') OR
                (Fields[J].FieldName = 'SHOP_IMAGE_LARGE') THEN
                BEGIN
                  IF J > 1 THEN
                    Buffer := Buffer + Delimiter;
                  Buffer := Buffer + Fields[J].FieldName;
                END;

            IF DosZeichenSatz THEN
              CharToOEM(PChar(Buffer), @Buffer[1]);

            Stream.Write(PChar(Buffer)^, Length(Buffer));
          END;

        First;
        WHILE NOT EOF DO
          BEGIN
            Buffer := #13#10;
            FOR J := 0 TO Fields.Count - 1 DO
              IF (Fields[J].FieldName = 'MATCHCODE') OR
                (Fields[J].FieldName = 'WARENGRUPPE') OR
                (Fields[J].FieldName = 'KURZNAME') OR
                (Fields[J].FieldName = 'LANGNAME') OR
                (Fields[J].FieldName = 'EK_PREIS') OR
                (Fields[J].FieldName = 'SHOP_IMAGE_LARGE') THEN
                BEGIN
                  FieldDesc := Fields[J];
                  IF J > 1 THEN
                    Buffer := Buffer + Delimiter;

                  IF NOT (FieldDesc.DataType IN [ftInteger, ftSmallInt, ftFloat,
                    ftAutoInc, ftCurrency, ftLargeInt,
                      ftBCD]) AND NOT FieldDesc.IsNull THEN
                    BEGIN
                      IF (Fields[J].FieldName = 'SHOP_IMAGE_LARGE') THEN
                        BEGIN
                          ArtikelBilderQuery.Close;
                          ArtikelBilderQuery.Sql.Clear;
                          RecID := Dataset.FieldByName('REC_ID').AsInteger;
                          ArtikelBilderQuery.Sql.Text :=
                            'SELECT * FROM ARTIKEL_BILDER WHERE ' +
                            'ARTIKEL_ID=' + IntToStr(RecID);
                          ArtikelBilderQuery.Open;
                          S := '';
                          IF (ArtikelBilderQuery.RecordCount > 0) THEN
                            S :=
                              ExtractFileName(ArtikelBilderQuery.FieldByName('BILD').AsString);
                        END
                      ELSE
                        S := FieldDesc.AsString;

                      IF TextInHochKomma THEN
                        Buffer := Buffer + AnsiQuotedStr(Uniquery.StringToSql(S),
                          '"')
                      ELSE
                        Buffer := Buffer + Uniquery.StringToSql(S);
                    END
                  ELSE
                    Buffer := Buffer + FieldDesc.AsString;
                END;
            IF DosZeichenSatz THEN
              CharToOEM(PChar(Buffer), @Buffer[1]);

            (*
                    try
                      UniQuery.Close;
                      UniQuery.Sql.Add('select * from Artikel_bilder');
                      UniQuery.Sql.Add('where Artikel_ID='+IntToStr(Dataset.FieldByName('Rec_ID').AsInteger));
                      UniQuery.Open;
                      IF UniQuery.RecordCount > 0 THEN
                        Buffer := Buffer + ExtractFileName(UniQuery.FieldByName('Bild').AsString);
                      UniQuery.Close;
                    except
                    end;
            *)
            Stream.Write(PChar(Buffer)^, Length(Buffer));
            inc(i);
            Next;
          END;
      END;
  FINALLY
    Screen.Cursor := oldC;
  END;
END;

FUNCTION TDM1.ObjektTransfer(Nr : Integer) : Boolean;
VAR
  QuellName, ZielName : STRING;
  ZielVerzeichnis : STRING;
VAR
  Error : Integer;
VAR
  NewID : Integer;
VAR
  Msg, Action : STRING;
  Data, Par : STRING;

BEGIN
  Screen.Cursor := crHourGlass;

  UniQuery2.Close;
  UniQuery2.Sql.Text := 'Select * from tbl_Obj where ObjNR = "' + inttoStr(Nr) +
    '"';
  UniQuery2.Open;
  Data := 'action=Objekt_Insert';
  Data := Data + '&objnr=' + UniQuery2.FieldByName('Objnr').AsString;
  Data := Data + '&wt=' + UniQuery2.FieldByName('Werbetext').AsString;
  Par := Data;
  Data := '';

  IF NOT ShopTransForm.UpdateData(Error, NewID, Msg, Action, Data, Par) THEN
    ;

  UniQuery2.Close;
  UniQuery2.Sql.Text := 'Select * from Obj_bilder where ObjNR = "' + inttoStr(Nr)
    + '"';
  UniQuery2.Open;

  IdFTP1.Username := ReadString('SHOP', 'FTP_BENUTZER', '');
  IdFTP1.Password := ReadString('SHOP', 'FTP_PASSWORT', '');
  IdFTP1.Host := ReadString('SHOP', 'FTP_SERVER', '');

  TRY
    IdFTP1.Connect;
  EXCEPT
    MessageBeep(MB_ICONHAND);
    ShowMessage('Die verbindung zum FTP Server(' + IdFTP1.Host +
      ') ist nicht möglich.'^M^M +
      Format('Fehler: %d', [GetLastError]));
    Screen.Cursor := crDefault;
    Exit;
  END;

  WHILE NOT UniQuery2.Eof DO
    BEGIN
      Zielverzeichnis := '/' + ReadString('SHOP', 'FTP_VERZEICHNIS', '') + '/';
      QuellName := Bildname(UniQuery2.FieldByName('Bild').AsString, 'Large');
      ZielName := Zielverzeichnis + ExtractFileName(QuellName);
      TRY
        IdFTP1.Put(QuellName, ZielName, FALSE);
      EXCEPT
        showmessage('Fehler Quelle:' + Quellname + ' Ziel:' + Zielname);
      END;
      UniQuery2.Next;
    END;
  IdFTP1.Disconnect;

  Screen.Cursor := crDefault;
END;

FUNCTION TDM1.Bildname(s, Typ : STRING) : STRING;
BEGIN

  Result := lowercase(ReadString('MAIN', 'STDBILDPFAD', '') + '\' + Copy(S, 1,
    Pos('.', S) - 1) + '_' +
    Typ + '.jpg');
  exit;

  IF Uppercase(ExtractFileExt(S)) = '.PDF' THEN
    BEGIN
      Result := lowercase(ReadString('MAIN', 'STDBILDPFAD', '') + '\' + 'pdf.jpg');
      exit;
    END;
  IF uppercase(typ) = 'thumb' THEN
    BEGIN
      Result := lowercase(ReadString('MAIN', 'STDBILDPFAD', '') + '\' + Copy(S, 1,
        Pos('.', S) - 1) + '_' + Typ + '.jpg');
      exit;
    END;

  IF Pos('.', S) > 0 THEN
    Result := lowercase(ReadString('MAIN', 'STDBILDPFAD', '') + '\' + Copy(S, 1,
      Pos('.', S) - 1) + '_' + Typ + ExtractFileExt(S))
  ELSE
    Result := lowercase(ReadString('MAIN', 'STDBILDPFAD', '') + '\' + S + '_' +
      Typ + '.jpg')
END;

FUNCTION TDM1.SummeAdd(Quelle, Was : STRING) : STRING;
VAR
  s : STRING;

BEGIN
  WHILE (Was > '') DO
    BEGIN
      IF pos(',', Was) = 0 THEN
        s := was
      ELSE
        s := copy(was, 1, pos(',', Was) - 1);
      IF Pos(s, Quelle) = 0 THEN
        //    If not SummeEnthalten(Quelle,s) then
        IF Quelle = '' THEN
          Quelle := Quelle + S
        ELSE
          Quelle := Quelle + ',' + S;
      IF pos(',', Was) > 0 THEN
        delete(was, 1, pos(',', Was))
      ELSE
        Was := '';
    END;
  result := Quelle;
END;


//------------------------------------------------------------------------------

PROCEDURE TDM1.AddLog(Typ, s : STRING);
VAR
  d : STRING;
  f : textfile;
  i : integer;
BEGIN
  d := DM1.LogDir + 'GERA_' + Typ + 'LOG_' + formatdatetime('dd_mm_yyyy', now) + '.log';

  // Unix To Dos Konvertieren (NL -> NL,LF)
  IF length(s) > 0 THEN
    BEGIN
      i := 0;
      WHILE i < length(s) DO
        BEGIN
          inc(i);
          IF (i > 1) AND (s[i] = #10) AND (s[i - 1] <> #13) THEN
            BEGIN
              insert(#13, s, i);
              inc(i);
            END; //if
        END; //while
    END; //if

  IF NOT fileexists(d) THEN
    fileclose(filecreate(d));
  assignfile(F, D);
  append(f);
  TRY
    writeln(f, S);
  FINALLY
    closefile(f);
  END;

END;
// brMenueAlle (Mitarbeiter Aquise Webtransfer )

FUNCTION Berechtigung(was : integer) : Boolean;
BEGIN
  If (was = brBelegLayout) and (DM1.MitarbeiterName = 'Zimmert') then
    Begin
      result := true;
      exit;
    end;

  IF (DM1.MitarbeiterName = 'Graef') OR (DM1.MitarbeiterName = 'Administrator') OR (DM1.MitarbeiterName = 'Gera') THEN
    result := true
  ELSE
    result := false;

END;

FUNCTION DateiSpeichern(S:String;Name:String;Overwrite:boolean=false):Integer;
var
  alterwert : integer;
  Dateiname : String;
BEGIN
  IF DateiZugriff = ZugriffNULL THEN
    DateiZugriffSuchen;


  Case Dateizugriff of
    ZugriffDirekt :
      BEGIN
        IF NOT CopyFile(PAnsiChar(s),PAnsiChar(DM1.DokumentenDir+Name),NOT OverWrite) THEN
          ShowMessage('Datei '+S+' konnte nicht auf '+DM1.DokumentenDir+Name+' kopiert werden.');
      end;
    ZugriffFTP :
      BEGIN
        DM1.IdFTP1.Host := DokumentenFTPHost;
        DM1.IdFTP1.Username := DokumentenFTPName;
        DM1.IdFTP1.Password := DokumentenFTPPasswort;
        DM1.IdFTP1.Connect();
        DM1.IdFTP1.Put(S,name);
        DM1.IdFTP1.Disconnect;
      END;
  end;

(*

  if not dm1.db2.Connected then
    dm1.db2.Connect;
    dm1.tabuni2.sql.text := 'show variables like "max_allowed_packet"';
    dm1.tabuni2.Open;
    alterwert := dm1.tabuni2.fieldbyname('value').asinteger;
    dm1.tabuni2.sql.text := 'set max_allowed_packet = '+inttostr(GetFileSize(S)+1024);
    dm1.tabuni2.ExecSql;
*)    
    (*
    dm1.tabData.open;
    dm1.tabData.Append;
    dm1.tabDataData.LoadFromFile(S);
    dm1.tabData.Post;
    dm1.tabuni2.sql.text := 'select last_insert_id()';
    dm1.tabuni2.Open;
    result := dm1.tabuni2.fieldbyname('last_insert_id()').asinteger;
    dm1.tabuni2.sql.text := 'set max_allowed_packet = '+inttostr(alterwert);
    dm1.tabuni2.ExecSql;
    *)
(*
    dm1.sqlquery1.SQL.Text := 'insert into data  set data = :blob ';
    dm1.sqlquery1.ParamByName('blob').LoadFromFile(s,ftBlob);
    dm1.sqlquery1.ExecSQL;
*)
      (*
      dm1.tabData.ParamByName('datablob').LoadFromFile(OpenDialog1.FileName, ftBlob);
      dm1.tabData.sql.text := 'insert into data set data = :datablob';
      dm1.tabData.ExecSql;
      *)
end;

Procedure DateiLoeschen(S : String);
BEGIN
   DeleteFile(s);
end;

END.

