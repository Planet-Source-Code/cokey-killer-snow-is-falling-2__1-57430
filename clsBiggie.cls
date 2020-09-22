VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBiggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is the biggie class

Private mX As Double    'Biggies X
Private mY As Double    'Biggies Y
Private mOldx As Double 'Biggies Old X
Private mOldy As Double 'Biggies Old Y
Private mXSpeed As Double   'Biggies X Speed.....this doesnt get set, so could be removed. :)
Private mYSpeed As Double   'Biggies Y Speed
Private mBiggieSize As Long 'The size that our biggie will be
Private mDead As Boolean    'Are we dead yet?
Private mColor As Long      'What color is the biggie going to be

'All the below stuff is the usual property GETs/LETs that allow us to
'set our biggie up

Public Property Let BiggieColor(ByVal InColor As Long)
    mColor = InColor
End Property
Public Property Get BiggieColor() As Long
    BiggieColor = mColor
End Property
Public Property Let x(ByVal InX As Double)
    mX = InX
End Property
Public Property Let y(ByVal InY As Double)
    mY = InY
End Property
Public Property Get x() As Double
    x = mX
End Property
Public Property Get y() As Double
    y = mY
End Property
Public Property Let xSpeed(ByVal Inxspeed As Double)
    mXSpeed = Inxspeed
End Property
Public Property Let ySpeed(ByVal InySpeed As Double)
    mYSpeed = InySpeed
End Property
Public Property Get xSpeed() As Double
    xSpeed = mXSpeed
End Property
Public Property Get ySpeed() As Double
    ySpeed = mYSpeed
End Property
Public Property Let BiggieSize(ByVal InSize As Long)
    mBiggieSize = InSize
End Property
Public Property Get BiggieSize() As Long
    BiggieSize = mBiggieSize
End Property
Public Property Let Dead(ByVal InDead As Boolean)
    mDead = InDead
End Property
Public Property Get Dead() As Boolean
    Dead = mDead
End Property
Public Property Let OldX(ByVal InOldX As Double)
    mOldx = InOldX
End Property
Public Property Let OldY(ByVal InOldY As Double)
    mOldy = InOldY
End Property
Public Property Get OldX() As Double
    OldX = mOldx
End Property
Public Property Get OldY() As Double
   OldY = mOldy
End Property




