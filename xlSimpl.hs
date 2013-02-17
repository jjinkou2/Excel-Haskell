import System.Win32.Com 
import System.Win32.Com.Automation
--
-- createObjectExcel 
-- coming from Automation.hs and com.hs
--
xlWorkbookDefault, xlSaveChanges, xlDoNotSaveChanges :: Int
xlDoNotSaveChanges = 2
xlSaveChanges = 1
xlWorkbookDefault = 51

getText :: IDispatch a -> IO String
getText obj= obj # propertyGet_0 "Text"


fichierTest2 = "E:/Programmation/haskell/Com/qos.xls"
fichierTest = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"

main = coRun $ do 
    pExl <- getFileObject fichierTest2 "Excel.Application"
    workBook <- getObject fichierTest2 
    pExl # propertySet "DisplayAlerts" [inBool False]
    workSheets <- workBook #  propertyGet_0 "Worksheets"
    sheetSel <- workSheets # propertyGet_1 "Item"  "BIV"
    rng <- sheetSel # propertyGet_1 "Range" "C7"
    text <-  rng # getText 
    putStrLn text

    -- workBook #  propertySet_1 "Saved" (1::Int)
    workBook # method_1_0 "Close"  xlDoNotSaveChanges

