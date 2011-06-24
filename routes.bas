Option Compare Database
Option Explicit

Function route(target As String)
  Forms(objForm).Controls(objControl).SourceObject = target
End Function