Attribute VB_Name = "Declare"
Option Explicit
Public totaldr As Double, totalcr As Double
Public TBOpeningBal As Double
Public OpeningBal As Double
Public SuspenseAcc As String, REarningsAcc As String, PPAcc As String
Dim colClass As Collection
Public cname As String
Public Email As String
Public motto As String
Public paddress As String
Public town As String
Public rsNumbers() As String
Public MsgContent As String
Public CPhone As String
Public Phone As String
Public sserver As Integer
Public AuditUser As String
Public transdate As Date
Public PrincAmount As Double
Public IntrAmount As Double
Public BankAmt As Double
Public SharesAmt As Double
Public Refno As Double
Public GlAccNBal As String
Public AuditName As String
Public GlAccBalance As Double
Public EarliestTransDate As Date
Public DormancyPeriod As Long, Offset As Long
Public glaccno As String
Public GlAccName As String
Public Int999 As Double
Public Group As Boolean
Public GlCode As String
Public glidno As String
Public glmemno As String
Public glpayno As String
Public rs2 As Recordset
Public DSource As String
Public GLteller As String
Public Level As String
Public Authorize As String
Public Editing As Boolean
Public FromAccNo As Boolean
Public strNewMembers As String
'// declaration for the third label
Public GlAccNo1 As String
Public TransNo As String
Public glnamE2 As String
Public GlIdNo1 As String
Public GlMemNo1 As String
Public GlPayNo1 As String
Public Form_Filled As Boolean
Public SchemeCode As String
'//declaration for the fourth label
Public glnamE1 As String
Public GlAccNo2 As String
Public GlName3 As String
Public GlIdNo2 As String
Public GlMemNo2 As String
Public GlPayNo2 As String
'//declaration for the fourth label
Public GlAccNo3 As String
Public GlName4 As String
Public GlIdNo3 As String
Public GlMemNo3 As String
Public GlPayNo3 As String
Public ReportTitle As String
Public ErrorMessage As String
Public bookba As Currency
Public title As String
Public STRFORMULA As String
Public reportname As String
Public CurrRecord As String
Public rsLoanGuar As Recordset
Public loanbalance As Double
Public AddInterest As Boolean
Public SelectedGroup As String
Public SelectedRec As String
Public Continue As Boolean
Public SelectedDate As Date
Public SelectedCompany As String
Public strValue As String
Public NewRecord As Boolean
Public Login As Boolean
Public Report_Path As String
Public AverageInterest As Boolean
Public SelectedDsn As String
Public permision As String
Public SearchForm As String
Public SearchValue As String
Public formCallingImport As Form
Public ImportForm As String
Public rangeFrom As String
Public rangeTo As String
Public action As String
Public lngRet As Long
Public User As String
Public username As String
Public FoundString As String
Public strRequiredNo  As String
Public counter2 As Integer
Public contrl As Control
Public activatepermision As String
Public theDatabase As String
Public theField As String
Public permisionShown As Control
Public EncryptPass As String
Public parametersSet As Boolean
Public SuperUser As String
Public MemberRegPermision As String
Public valToEncrOrDecr As String
Public theButtonClickedOnMainForm As String
Public acceptmodify As Boolean
Public applicationpermision As String
Public guarantorspermision As String
Public SummaryWasVisible As Boolean
Public theColumns As String
Public databasesetpermision As String
Public rstRecordsImported As ADODB.Recordset
Public usergroupspermision As String
Public backuppermision As String
Public mycontainer As Frame
Public memstatementspermision As String
Public utilstatementpermision As String
Public sysuserspermision As String
Public loanguarantorspermision As String
Public transactionspermision As String
Public rejreasonpermision As String
Public utilguarantorspermision As String
Public loanendorsementpermision As String
Public changenopermision As String
Public clearmempermision As String
Public clearloanpermision As String
Public sharevarpermision As String
Public companysetpermision As String
Public banksetuppermision As String
Public endorsementpermision As String
Public periodtranpermision As String
Public utilguarpermision As String
Public parampermision As String
Public loantypespermision As String
Public archivedpermision As String
Public dormantpermision As String
Public monthlydedpermision As String
Public chequeentrypermision As String
Public savingspermision As String
Public deductionspermision As String
Public withdrawnpermision As String
Public benfundpermision As String
Public exporttoglpermision As String
Public cancelActionInvolvingRange As Boolean
Public calculatorpermision As String
Public dividendspermision As String
Public effectrepaypermision As String
Public contributionspermision As String
Public chkchequeentrypermisions As String
Public loanbalpermision As String
Public NextOfKinPermision As String
Public statementpermision As String
Public manuallychecked As String
Public databaseLocked As Boolean
Public percentageDifference As Integer
Public atStartupOfSystem As Boolean
Public theDatabaseLocked As String
Public reportType As String
Public formCallingExport As Form
Public theTextBoxOnRange As String
Public formCallingRangeSelector As Form
Public a As CRAXDDRT.Application
Public r As CRAXDDRT.Report
Public oSaccoMaster As New CSaccoData 'reference to the class
Public rsMembership
Public controlsCleared As Boolean
Public strSQL As String
Public cn As Connection
Public rs As Recordset
Public Rst As Recordset
Public Rst1 As Recordset
Public rst2 As Recordset
Public Rst3 As Recordset
Public Rst4 As Recordset
Public Rst5 As ADODB.Recordset
Public rst6 As Recordset
Public li As ListItem
Public Li2 As ListItem
Public SItem As Integer
Public sel As String
Public loan As Currency
Public prince As Currency
Public MaxAmount, initshares As Currency
Public Shares As Currency
Public LoanNo As String
Public compcodes As String
Public bankcodes As String
Public I As Long
Public mMemberNo As String
Public strFileName As String
Public strSignName As String
Public LtSRatio As Long
Public memPic As FileSystemObject
Public rate As Integer
Public name As String
Public DataB As String
Public Fmt As String
Public sql As String
Public memnum As String
Public MyBookMark
Public searchField As String
Dim recordfound As String
Public formCallingSearch As Form
Public dateDifference As Integer
Public counter As Integer
Public Const Cfmt = "###,###,###,###,###0.00"
Public Const DCfmt = "###,###,###,###,###0.00"
Public Const Dfmt = "mm-dd-YY"
Public FirstPay As Boolean
Public MyRecord As String
Public RecAmount As Double
Public mno As String
Public mName As String
Public vno As String
Public auditid As String
Public mysql As String
