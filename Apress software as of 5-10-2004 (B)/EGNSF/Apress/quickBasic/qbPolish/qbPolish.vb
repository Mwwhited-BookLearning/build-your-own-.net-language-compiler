Option Strict On

Imports System.Threading
Public Class qbPolish

    ' *********************************************************************
    ' *                                                                   *
    ' * qbPolish     quickBasicEngine: Polish instruction                 *
    ' *                                                                   *
    ' *                                                                   *
    ' * This class represents one instruction to our non-CLR Nutty        *
    ' * Professor machine: for more information see qbPolish.DOC.         *  
    ' *                                                                   *
    ' *********************************************************************

    ' ***** Shared *****
    Private Shared _INTsequence As Integer
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJop As qbOp.qbOp

    ' ***** Constants *****
    ' --- Class name
    Private Const CLASS_NAME As String = "qbPolish"
    ' --- Easter yEgg
    Private Const ABOUTINFO As String =
        "qbPolish" &
        vbNewLine & vbNewLine &
        "This class represents one instruction to our non-CLR Nutty Professor machine." &
        vbNewLine & vbNewLine &
        "This class was developed commencing on 6/04/2003 by" & vbNewLine & vbNewLine &
        "Edward G. Nilges" & vbNewLine &
        "spinoza1111@yahoo.COM" & vbNewLine &
        "http://members.screenz.com/edNilges"
    ' --- Inspection
    Private Const INSPECTION_USABLE As String =
        "The object must be usable"
    Private Const INSPECTION_OPVALID As String =
        "The operation code can’t be the Invalid enumerator value"
    Private Const INSPECTION_OPERAND As String =
        "The operand must be a .Net scalar"
    Private Const INSPECTION_STARTINDEXLENGTH As String =
        "The start index of the source code for the op must be greater than " &
        "0 AND the length must be greater than 0, " &
        "or else, both start index and length should be 0"

    ' ***** Object State *****
    Private Structure TYPstate
        Dim strName As String                ' Object instance name
        Dim booUsable As Boolean             ' True: object is usable
        Dim enuOpCode As qbOp.qbOp.ENUop     ' Operator
        Dim objOperand As Object             ' Operand
        Dim strComment As String             ' Commentary
        Dim intStartIndex As Integer         ' Source code token start index
        Dim intLength As Integer             ' Source code token length
    End Structure
    Private USRstate As TYPstate

    ' ***** Object constructor ****************************************

    Public Sub New()
        With USRstate
            .strName = "qbPolish" &
                       _OBJutilities.alignRight(CStr(Interlocked.Increment(_INTsequence)),
                                                4, "0")
            Dim strTestReport As String
            .booUsable = True : .booUsable = inspection_()
        End With
    End Sub

    ' ***** Public Procedures *****************************************

    ' -----------------------------------------------------------------
    ' Return information about the class
    '
    '
    Public Shared ReadOnly Property About As String
        Get
            Return (ABOUTINFO)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return name of class
    '
    '
    Public Shared ReadOnly Property ClassName As String
        Get
            Return (CLASS_NAME)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return and change commentary
    '
    '
    Public Property Comment As String
        Get
            If Not checkUsable_("Comment Get") Then Return ""
            Return (USRstate.strComment)
        End Get
        Set(ByVal strNewValue As String)
            If Not checkUsable_("Comment Set") Then Return
            USRstate.strComment = strNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Object disposer
    '
    '
    Public Function dispose() As Boolean
        USRstate.booUsable = False
    End Function

    ' ------------------------------------------------------------------
    ' Inspect the object instance
    '
    '
    Public Function inspect(ByRef strReport As String) As Boolean
        With USRstate
            Dim booInspection As Boolean = True
            strReport = "Inspection of qbPolish instance " &
                        _OBJutilities.enquote(.strName) & " " &
                        "at " & Now
            If _OBJutilities.inspectionAppend(strReport,
                                              INSPECTION_USABLE,
                                              .booUsable,
                                              booInspection) Then
                _OBJutilities.inspectionAppend(strReport,
                                                INSPECTION_OPVALID,
                                                .enuOpCode <> qbOp.qbOp.ENUop.opInvalid,
                                                booInspection)
                _OBJutilities.inspectionAppend(strReport,
                                                INSPECTION_STARTINDEXLENGTH,
                                                .intStartIndex > 0 AndAlso .intLength > 0 _
                                                OrElse
                                                .intStartIndex = 0 AndAlso .intLength = 0,
                                                booInspection,
                                                "Start index is " & .intStartIndex & ": " &
                                                "length is " & .intLength)
            End If
            If Not booInspection Then
                .booUsable = False : Return (False)
            End If
            Return (True)
        End With
    End Function

    ' ------------------------------------------------------------------
    ' Make the object not usable
    '
    '
    Public Function mkUnusable() As Boolean
        USRstate.booUsable = False
    End Function

    ' ------------------------------------------------------------------
    ' Return and change the name of the object
    '
    '
    Public Property Name() As String
        Get
            Return (USRstate.strName)
        End Get
        Set(ByVal strNewValue As String)
            USRstate.strName = strNewValue
        End Set
    End Property

    ' ------------------------------------------------------------------
    ' Convert the object to eXtended Markup Language
    '
    '
    ' --- Default commenting (full commenting)
    Public Overloads Function object2XML() As String
        Return (object2XML(True, True))
    End Function
    ' --- Allows suppression of the header comment
    Public Overloads Function object2XML(ByVal booHeaderComment As Boolean) _
           As String
        Return (object2XML(booHeaderComment, True))
    End Function
    ' --- Allows suppression of both comments   
    Public Overloads Function object2XML(ByVal booHeaderComment As Boolean,
                                         ByVal booLineComments As Boolean) _
           As String
        With USRstate
            Return (_OBJutilities.objectInfo2XML(Me.ClassName,
                                   Me.About,
                                   booHeaderComment,
                                   booLineComments,
                                   "booUsable",
                                   "Indicates usability of object",
                                   CStr(.booUsable),
                                   "strName",
                                   "Names the object instance",
                                   .strName,
                                   "enuOpCode",
                                   "Identifies the op code",
                                   .enuOpCode.ToString,
                                   "intStartIndex",
                                   "Identifies the first token responsible for this op code",
                                   CStr(.intStartIndex),
                                   "intLength",
                                   "Identifies how many tokens are responsible for this op code",
                                   CStr(.intLength),
                                   "objOperand",
                                   "Identifies the operand (if any)",
                                   _OBJutilities.object2String(.objOperand),
                                   "strComment",
                                   "Comments the opcode",
                                   .strComment))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return and assign operation code
    '
    '
    Public Property Opcode() As qbOp.qbOp.ENUop
        Get
            If Not checkUsable_("Opcode Get") Then Return (qbOp.qbOp.ENUop.opInvalid)
            Return USRstate.enuOpCode
        End Get
        Set(ByVal enuOpCode As qbOp.qbOp.ENUop)
            If Not checkUsable_("Opcode Set") Then Return
            USRstate.enuOpCode = enuOpCode
            Return
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return and assign operation code
    '
    '
    Public Function opcodeFromString(ByVal strOpcode As String) As Boolean
        If Not checkUsable_("opcodeFromString") Then Return (False)
        USRstate.enuOpCode = qbOp.qbOp.opCodeFromString(strOpcode)
    End Function

    ' -----------------------------------------------------------------
    ' Return description of operation code
    '
    '
    Public Function opcodeToDescription() As String
        If Not checkUsable_("opcodeToDescription") Then Return ("")
        Return (qbOp.qbOp.opCodeToDescription(USRstate.enuOpCode))
    End Function

    ' -----------------------------------------------------------------
    ' Return string name of operation code
    '
    '
    Public Function opcodeToString() As String
        If Not checkUsable_("opcodeToString") Then
            Return ("")
        End If
        Return (qbOp.qbOp.opCodeToString(USRstate.enuOpCode))
    End Function

    ' -----------------------------------------------------------------
    ' Return and set the operand
    '
    '
    Public Property Operand() As Object
        Get
            If Not checkUsable_("Operand Get") Then
                Return (Nothing)
            End If
            Return (USRstate.objOperand)
        End Get
        Set(ByVal objNewValue As Object)
            If Not checkUsable_("Operand Set") Then
                Return
            End If
            USRstate.objOperand = objNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return and change length of associated source code in tokens
    '
    '
    Public Property TokenLength() As Integer
        Get
            If Not checkUsable_("TokenLength Get") Then Return (0)
            Return (USRstate.intLength)
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("TokenLength Set") Then Return
            If intNewValue < 0 Then
                errorHandler_("Invalid source length " & intNewValue,
                              "TokenLength Set",
                              "Not changing object state")
                Return
            End If
            USRstate.intLength = intNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return and change start index of associated source code
    '
    '
    Public Property TokenStartIndex() As Integer
        Get
            If Not checkUsable_("TokenStartIndex Get") Then Return (0)
            Return (USRstate.intStartIndex)
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("TokenStartIndex Set") Then Return
            If intNewValue < 1 Then
                errorHandler_("Invalid source start index " & intNewValue,
                              "SourceStartIndex Set",
                              "Not changing object state")
                Return
            End If
            USRstate.intStartIndex = intNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Convert to string
    '
    '
    Public Overrides Function toString() As String
        If Not checkUsable_("toString") Then Return ("")
        With Me
            Dim strOperand As String
            If Not _OBJop.isJumpOp(.Opcode) _
               AndAlso
               Not IsNumeric(strOperand) _
               AndAlso
               Not _OBJutilities.isQuoted(strOperand) _
               AndAlso
               Not (.Operand Is Nothing) Then
                strOperand = _OBJutilities.object2String(.Operand)
            ElseIf (TypeOf .Operand Is qbVariable.qbVariable) Then
                strOperand = CType(.Operand, qbVariable.qbVariable).toString
            ElseIf Not (.Operand Is Nothing) Then
                strOperand = CStr(.Operand)
            End If
            Return (.opcodeToString & " " &
                    strOperand &
                    CStr(IIf(Trim(.Comment) = "", "", ": ")) &
                    .Comment)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return object usability
    '
    '
    Public ReadOnly Property Usable() As Boolean
        Get
            Return (USRstate.booUsable)
        End Get
    End Property

    ' ***** Private Procedures *****************************************

    ' -----------------------------------------------------------------
    ' Check usability
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String) As Boolean
        If Not USRstate.booUsable Then
            errorHandler_("Object instance is not usable",
                          strProcedure,
                          "Returning a default value: making no change to " &
                          "the state of the object")
        End If
        Return (True)
    End Function

    ' ------------------------------------------------------------------
    ' Interface to the error handler
    '
    '
    Private Sub errorHandler_(ByVal strMessage As String,
                              ByVal strProcedure As String,
                              ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage,
                                   Me.Name,
                                   strProcedure,
                                   strHelp)
    End Sub

    ' -------------------------------------------------------------------
    ' Internal inspection
    '
    '
    Private Function inspection_() As Boolean
        Dim strReport As String
        If Not Me.inspect(strReport) Then
            errorHandler_("Internal inspection failed: " &
                          vbNewLine & vbNewLine &
                          strReport,
                          "inspection_",
                          "Object has been marked unusable")
            Return (False)
        End If
        Return (True)
    End Function

End Class
