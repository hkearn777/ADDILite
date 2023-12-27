Public Class fieldInfo
  Private _level As String
  Private _fieldName As String
  Private _redefines As String
  Private _picture As String
  Private _usage As String
  Private _startPos As Integer
  Private _endPos As Integer
  Private _length As Integer
  Private _occursMinimumTimes As Integer
  Private _occursMaximumTimes As Integer
  Private _dependingOn As String
  Private _parent As Integer
  Private _redefField As Integer

  Public Property Level As String
    Get
      Return _level
    End Get
    Set(value As String)
      _level = value
    End Set
  End Property

  Public Property FieldName As String
    Get
      Return _fieldName
    End Get
    Set(value As String)
      _fieldName = value
    End Set
  End Property

  Public Property Redefines As String
    Get
      Return _redefines
    End Get
    Set(value As String)
      _redefines = value
    End Set
  End Property

  Public Property Picture As String
    Get
      Return _picture
    End Get
    Set(value As String)
      _picture = value
    End Set
  End Property

  Public Property Usage As String
    Get
      Return _usage
    End Get
    Set(value As String)
      _usage = value
    End Set
  End Property

  Public Property StartPos As Integer
    Get
      Return _startPos
    End Get
    Set(value As Integer)
      _startPos = value
    End Set
  End Property

  Public Property EndPos As Integer
    Get
      Return _endPos
    End Get
    Set(value As Integer)
      _endPos = value
    End Set
  End Property

  Public Property Length As Integer
    Get
      Return _length
    End Get
    Set(value As Integer)
      _length = value
    End Set
  End Property

  Public Property OccursMinimumTimes As Integer
    Get
      Return _occursMinimumTimes
    End Get
    Set(value As Integer)
      _occursMinimumTimes = value
    End Set
  End Property

  Public Property OccursMaximumTimes As Integer
    Get
      Return _occursMaximumTimes
    End Get
    Set(value As Integer)
      _occursMaximumTimes = value
    End Set
  End Property

  Public Property DependingOn As String
    Get
      Return _dependingOn
    End Get
    Set(value As String)
      _dependingOn = value
    End Set
  End Property
  Public Property Parent As Integer
    Get
      Return _parent
    End Get
    Set(value As Integer)
      _parent = value
    End Set
  End Property

  Public Property RedefField As Integer
    Get
      Return _redefField
    End Get
    Set(value As Integer)
      _redefField = value
    End Set
  End Property

  Public Sub New()

  End Sub
  Public Sub New(ByVal _level As String,
                   ByVal _fieldName As String,
                   ByVal _redefines As String,
                   ByVal _picture As String,
                   ByVal _usage As String,
                   ByVal _startPos As Integer,
                   ByVal _endPos As Integer,
                   ByVal _length As Integer,
                   ByVal _occursMinimumTimes As Integer,
                   ByVal _occursMaximumTimes As Integer,
                   ByVal _DependingOn As String,
                   ByVal _parent As Integer,
                   ByVal _redefField As Integer)
    Level = _level
    FieldName = _fieldName
    Redefines = _redefines
    Picture = _picture
    Usage = _usage
    StartPos = _startPos
    EndPos = _endPos
    Length = _length
    OccursMinimumTimes = _occursMinimumTimes
    OccursMaximumTimes = _occursMaximumTimes
    DependingOn = _DependingOn
    Parent = _parent
    RedefField = _redefField
  End Sub

End Class
