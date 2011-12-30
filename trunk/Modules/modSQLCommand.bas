Attribute VB_Name = "modSQLCommand"
Option Explicit

Global Const g_SQL_Client = "SHAPE {SELECT * FROM CLIENT}" _
                      & " APPEND({SELECT * FROM CLIENT_CUSTOMER WHERE cCode=? AND cCompanyID = ?} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsCHeader," _
                      & " ({SELECT * FROM CLIENT_SUPPLIER WHERE (cCode=?)  AND (cCompanyID = ?)} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsSHeader, " _
                      & " ({SELECT * FROM CLIENT_CUSTOMER_ADD WHERE (cCode=?) AND (cCompanyID = ?)} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsAddress, " _
                      & " ({SELECT * FROM CLIENT_CUSTOMER_CL WHERE cType = 'MS' AND (cCode=?) AND (cCompanyID = ?)} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsMS, " _
                      & " ({SELECT * FROM CLIENT_CUSTOMER_CL WHERE cType = 'PD' AND (cCode=?) AND (cCompanyID = ?)} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsPD, " _
                      & " ({SELECT * FROM CLIENT_SUPPLIER_ITEM WHERE (cCode=?) AND (cCompanyID = ?)} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsSProduct, " _
                      & " ({SELECT * FROM CLIENT_SUPPLIER_ADD WHERE (cCode=?) AND (cCompanyID = ?)} RELATE cCode TO PARAMETER 0, cCompanyID TO PARAMETER 1) AS rsSAddress"


