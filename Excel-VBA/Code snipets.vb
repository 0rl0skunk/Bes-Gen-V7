PlankopfFactory.create( _
    Projekt:=, _
    GezeichnetPerson:=, _
    GezeichnetDatum:=, _
    GeprüftPerson:=, _
    GeprüftDatum:=, _
    Gebäude:=, _
    Gebäudeteil:=, _
    Geschoss:=, _
    Format:=, _
    Masstab:=, _
    Stand:=, _
    Klartext:=, _
    KlartextGebäude:=, _
    KlartextGeschoss:=, _
    TinLineID:=, _
    SkipValidation:= _
)

ValidInputs()
    dim ErrorSource as string
    ErrorSource="Plankopf > ValidInputs"
    if len(Inputs.ID)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.IDTinLine)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    '--- Zeichner ---
    if len(Inputs.GezeichnetPerson)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.GezeichnetDatum)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.GeprüftPerson)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.GeprüftDatum)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    '--- Gebäude ---
    if len(Inputs.Gebäude)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.Gebäudeteil)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.Geschoss)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    '--- Planbezeichnung ---
    if len(Inputs.Klartext)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.KlartextGebäude)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.KlartextGeschoss)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.Planüberschrift)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    '--- Layout ---
    if len(Inputs.LayoutName)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.LayoutGrösse)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.LayoutMasstab)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.LayoutPlanstand)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    '--- File Path ---
    if len(Inputs.PDFFileNaInputs.me)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.DWGFileNaInputs.me)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.DWGFilePath)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.XMLFileNaInputs.me)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 
    if len(Inputs.XMLFilePath)=0 then err.raise 1,ErrorSource,"no 'INPUT' given" 

    Inputs

    inputs.ID=ID
    inputs.IDTinLine=IDTinLine
    inputs.GezeichnetPerson=GezeichnetPerson
    inputs.GezeichnetDatum=GezeichnetDatum
    inputs.GeprüftPerson=GeprüftPerson
    inputs.GeprüftDatum=GeprüftDatum
    inputs.Gebäude=Gebäude
    inputs.Gebäudeteil=Gebäudeteil
    inputs.Geschoss=Geschoss
    inputs.Klartext=Klartext
    inputs.KlartextGebäude=KlartextGebäude
    inputs.KlartextGeschoss=KlartextGeschoss
    inputs.Planüberschrift=Planüberschrift
    inputs.LayoutGrösse=LayoutGrösse
    inputs.LayoutMasstab=LayoutMasstab
    inputs.LayoutPlanstand=LayoutPlanstand