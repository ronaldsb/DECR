Attribute VB_Name = "Module"
    
    
    'Declaración de la variables globales
    
    
    Public Datos(), Frame(), Posicion, FrameU As Variant
    
    
    Public SismoM(7), SismoC(7), SismoT(7), CombosM(12, 12), CombosC(12, 12), CombosT(12, 12) As Variant
    
    
    Public Dir, Protegido, Archivo, Hechos, Proyectos As String
    
    
    Public Pn2(18), Mn2(18), fPn2(18), fMn2(18), Pnc2(18), Mnc2(18), Pnp2(18), Mnp2(18) As Variant
    
    
    Public Pu, puntos, Mmax As Variant
    
      
    Sub Main()
  
    
    Protegido = "N"  '  ("S"   =   Sí) ; ("N"   =   NO)
    
    
    FDialog.Visible = True
    
   
    End Sub
