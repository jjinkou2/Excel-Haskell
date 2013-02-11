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

getActiveWB :: IDispatch a -> IO (IDispatch ()) 
getActiveWB obj = obj # propertyGet_0 "ActiveWorkBook"
--getRange  :: Variant b => String -> IDispatch a -> IO b 
getText :: IDispatch a -> IO String
getText obj= obj # propertyGet_0 "Text"


iidIDispatch_unsafe  = mkIID "{00020400-0000-0000-C000-000000000046}"

createObjExl :: IO (IDispatch ()) 
createObjExl = do
    clsidExcel <- clsidFromProgID "Excel.Application"
    pExl <- coCreateInstance clsidExcel  Nothing LocalProcess iidIDispatch_unsafe
    return pExl


fichierTest2 = "E:/Programmation/haskell/Com/qos1.xls"
fichierTest = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"

main = coRun $ do 
    pExl <- createObjExl 
    workBooks <- pExl #  propertyGet_0 "Workbooks" 
    workBook <- workBooks #  propertyGet_1 "Open" fichierTest
    workSheets <- workBook #  propertyGet_0 "Worksheets"
    sheetSel <- workSheets # propertyGet_1 "Item" (1::Int) :: IO (IDispatch ())
    rng <- sheetSel # propertyGet_1 "Range" "C3"
    text <-  rng # getText 
    putStrLn text

    workBook #  propertySet_1 "Saved" (1::Int)
    workBooks # method_1_0 "Close"  xlDoNotSaveChanges
    pExl # method_0_0 "Quit"

    mapM release [rng,sheetSel, workSheets,workBook, workBooks, pExl]
    --mapM release [workBook,workBooks, pExl]
    --mapM release [workBooks, pExl]
