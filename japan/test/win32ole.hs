{-# LANGUAGE ForeignFunctionInterface #-}
-- ghc --make win32ole.hs  winole.c -lole32 -loleaut32 -luuid  -o ole
module Main where

import Foreign.Ptr           (Ptr)
import Foreign.C.String      (CWString, newCWString, withCWString, peekCWString)
import Foreign.C.Types       (CLong, CInt)
import Foreign.Marshal.Alloc (free)
import Control.Applicative   ((<$>))

data IDispatch = IDispatch

main = do
    cOleInitialize 0
    pExl          <- instanceNew        "Excel.Application"
    cVersString   <- readProperty pExl  "Version"
    (("Version:"++) <$> peekCWString cVersString) >>= putStrLn
    cSysFreeString cVersString

    workBooks     <- propertyGet_S pExl "Workbooks"
    cFullFileName <- getFullPathName    "sample2.xls"
    workBooksOpen workBooks cFullFileName
    free cFullFileName

    activeWBook   <- propertyGet_S   pExl        "ActiveWorkbook"
    workSheets    <- propertyGet_S   activeWBook "Worksheets"
    sheet         <- propertyGet_S_N workSheets  "Item"  2    
    cell          <- propertyGet_S_S sheet       "Range" "C1"
    propertyPut_S_S cell "Value" "tarte"                    
    cwString   <- readProperty cell "Value"                
    (("Version:"++) <$> peekCWString cwString) >>= putStrLn
    cSysFreeString cwString

    mapM_ method_S       [(activeWBook,"Save"),(workBooks,"Close"),(pExl, "Quit")]
    mapM_ cReleaseObject [cell, sheet, workSheets, activeWBook, workBooks, pExl]
    cOleUninitialize

instanceNew :: String -> IO (Ptr IDispatch)
instanceNew name = withCWString name cInstanceNew

readProperty :: (Ptr IDispatch) -> String -> IO (CWString)
readProperty pDisp name = withCWString name (cReadProperty pDisp)

propertyGet_S :: (Ptr IDispatch) -> String -> IO (Ptr IDispatch)
propertyGet_S pDisp name = withCWString name (cPropertyGet_S pDisp)

getFullPathName :: String -> IO (CWString)
getFullPathName fName =  withCWString fName cgetFullPathName

propertyGet_S_N :: (Ptr IDispatch) -> String -> CLong -> IO (Ptr IDispatch)
propertyGet_S_N pDisp name n = withCWString name $ \x -> cPropertyGet_S_N pDisp x  n

propertyGet_S_S :: (Ptr IDispatch) -> String -> String -> IO (Ptr IDispatch)
propertyGet_S_S pDisp command param = withCWString command (\x ->withCWString param (cPropertyGet_S_S pDisp x))

propertyPut_S_S :: (Ptr IDispatch) -> String -> String -> IO ()
propertyPut_S_S pDisp name value = withCWString name (\x ->withCWString value ( cPropertyPut_S_S pDisp x))

method_S :: ((Ptr IDispatch), String) -> IO ()
method_S (pDisp, name) = withCWString name (cMethod_S pDisp)

workBooksOpen  :: (Ptr IDispatch) -> CWString -> IO ()
workBooksOpen pDisp fileName =  withCWString "Open" (\x -> cMethod_S_S pDisp x fileName)

-- C
foreign import ccall   "InstanceNew"            cInstanceNew       :: CWString -> IO (Ptr IDispatch)
foreign import ccall   "getFullPathName"        cgetFullPathName   :: CWString -> IO CWString
foreign import ccall   "PropertyGet_S"          cPropertyGet_S     :: (Ptr IDispatch) -> CWString -> IO (Ptr IDispatch)
foreign import ccall   "PropertyGet_S_S"        cPropertyGet_S_S   :: (Ptr IDispatch) -> CWString -> CWString -> IO (Ptr IDispatch)
foreign import ccall   "PropertyGet_S_N"        cPropertyGet_S_N   :: (Ptr IDispatch) -> CWString -> CLong -> IO (Ptr IDispatch)
foreign import ccall   "PropertyPut_S_S"        cPropertyPut_S_S   :: (Ptr IDispatch) -> CWString -> CWString -> IO ()
foreign import ccall   "ReadProperty"           cReadProperty      :: (Ptr IDispatch) -> CWString -> IO CWString
foreign import ccall   "Method_S_S"             cMethod_S_S        :: (Ptr IDispatch) -> CWString -> CWString -> IO ()
foreign import ccall   "Method_S"               cMethod_S          :: (Ptr IDispatch) -> CWString -> IO ()
foreign import ccall   "ReleaseObject"          cReleaseObject     :: (Ptr IDispatch) -> IO ()
foreign import ccall   "stdlib.h free"          cfree              :: CWString -> IO ()
foreign import ccall   "stdlib.h free"          cDispatchFree      :: (Ptr IDispatch) -> IO ()
foreign import stdcall "windows.h SysFreeString"  cSysFreeString   :: CWString -> IO ()
foreign import stdcall "ole2.h OleInitialize"     cOleInitialize   :: CInt -> IO ()
foreign import stdcall "ole2.h OleUninitialize"   cOleUninitialize :: IO ()

