
import ExcelCom 


fichierTest = "C:/Users/lyup7323/Developpement/Haskell/Com/qos.xls"
fichierTest1 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"


main = coRun $ do   
    pExl <- createObjExl
    workBooks <- pExl # getWorkbooks
    pExl # propertySet "DisplayAlerts" [inBool False]
    workBook <- workBooks # openWorkBooks fichierTest1
    workSheets <- workBook # getWSheets'
    sheetSel <- workSheets # propertyGet_1 "Item" (1::Int)

    rng <- sheetSel # propertyGet_1 "Range" "C3"
    text <-  rng # getText 
    putStrLn text
    


    workBooks # method_1_0 "Close" xlDoNotSaveChanges
    pExl # method_0_0 "Quit"
    
