{-# LANGUAGE ForeignFunctionInterface #-}
-- ghc --make win32ole.hs  winole.c -lole32 -loleaut32 -luuid  -o ole
module Main where

import Foreign.Ptr           (Ptr)
import Foreign.ForeignPtr    (ForeignPtr (..))
import Foreign.C.String      (CWString, newCWString, withCWString, peekCWString)
import Foreign.C.Types       (CLong(..), CInt(..))
import Foreign.Marshal.Alloc (free)
import Memory

fichierTest3 = "E:/Programmation/haskell/Com/qos.xls"

main = do
    cOleInitialize 0
    pExl          <- instanceNew        "Excel.Application"

    workBooks     <- propertyGet_S pExl "Workbooks"
    --cFullFileName <- getFullPathName    "qos.xls"
    cFullFileName <- newCWString    fichierTest3
    workBooksOpen workBooks cFullFileName
    free cFullFileName

    activeWBook   <- propertyGet_S   pExl        "ActiveWorkbook"
    workSheets    <- propertyGet_S   activeWBook "Worksheets"
    sheet         <- propertyGet_S_S workSheets  "Item"  "BIV"    
    cell          <- propertyGet_S_S sheet       "Range" "C7"
    -- propertyPut_S_S cell "Value" "toto"                     
    cwString      <- readProperty cell "Value" 
    peekCWString cwString >>= putStrLn 
    cSysFreeString cwString

    mapM_ method_S       [(activeWBook,"Save"),(workBooks,"Close"),(pExl, "Quit")]
    mapM_ cReleaseObject [cell, sheet, workSheets, activeWBook, workBooks, pExl]
    cOleUninitialize

--data IDispatch a = IDispatch a
-- --------------------------------------------------
-- 
-- interface IDispatch a
-- 
-- --------------------------------------------------
data IDispatch_ a = IDispatch__
                      
type IDispatch a = IUnknown (IDispatch_ a)

newtype IUnknown_ a  = Unknown  (ForeignPtr ())
type IUnknown  a  = IUnknown_ a

--type HRESULT = Int32
-- --------------------------------------------------
-- 
-- interface EnumVariant
-- 
-- --------------------------------------------------
data EnumVARIANT a      = EnumVARIANT
type IEnumVARIANT a     = IUnknown (EnumVARIANT a)
--iidIEnumVARIANT :: IID (IEnumVARIANT ())
--iidIEnumVARIANT = mkIID "{00020404-0000-0000-C000-000000000046}"

-- --------------------------------------------------
-- 
-- interface Com from Com.hs
-- 
-- --------------------------------------------------
-- | @stringToGUID "{00000000-0000-0000-C000-0000 0000 0046}"@ translates the
-- COM string representation for GUIDs into an actual 'GUID' value.
{-stringToGUID :: String -> IO GUID
stringToGUID str =
   stackWideString str $ \xstr -> do
   pg <- coAlloc sizeofGUID
   primStringToGUID xstr (castPtr pg)
   unmarshallGUID True pg
-}
sizeofGUID  :: Word32
sizeofGUID  = 16
-- --------------------------------------------------
-- 
-- Helpers/Converters
-- 
-- --------------------------------------------------
instanceNew :: String -> IO (IDispatch a)
instanceNew name = withCWString name cInstanceNew

readProperty :: (IDispatch a) -> String -> IO (CWString)
readProperty pDisp name = withCWString name (cReadProperty pDisp)

propertyGet_S :: (IDispatch a) -> String -> IO (IDispatch a)
propertyGet_S pDisp name = withCWString name (cPropertyGet_S pDisp)

getFullPathName :: String -> IO (CWString)
getFullPathName fName =  withCWString fName cgetFullPathName

propertyGet_S_N :: (IDispatch a) -> String -> CLong -> IO (IDispatch a)
propertyGet_S_N pDisp name n = withCWString name $ \x -> cPropertyGet_S_N pDisp x  n

propertyGet_S_S :: (IDispatch a) -> String -> String -> IO (IDispatch a)
propertyGet_S_S pDisp command param = withCWString command (\x ->withCWString param (cPropertyGet_S_S pDisp x))

propertyPut_S_S :: (IDispatch a) -> String -> String -> IO ()
propertyPut_S_S pDisp name value = withCWString name (\x ->withCWString value ( cPropertyPut_S_S pDisp x))

method_S :: ((IDispatch a), String) -> IO ()
method_S (pDisp, name) = withCWString name (cMethod_S pDisp)

workBooksOpen  :: (IDispatch a) -> CWString -> IO ()
workBooksOpen pDisp fileName =  withCWString "Open" (\x -> cMethod_S_S pDisp x fileName)

-- --------------------------------------------------
-- 
-- C interface
-- 
-- --------------------------------------------------
foreign import ccall   "InstanceNew"            cInstanceNew       :: CWString -> IO (IDispatch a)
foreign import ccall   "getFullPathName"        cgetFullPathName   :: CWString -> IO CWString
foreign import ccall   "PropertyGet_S"          cPropertyGet_S     :: (IDispatch a) -> CWString -> IO (IDispatch a)
foreign import ccall   "PropertyGet_S_S"        cPropertyGet_S_S   :: (IDispatch a) -> CWString -> CWString -> IO (IDispatch a)
foreign import ccall   "PropertyGet_S_N"        cPropertyGet_S_N   :: (IDispatch a) -> CWString -> CLong -> IO (IDispatch a)
foreign import ccall   "PropertyPut_S_S"        cPropertyPut_S_S   :: (IDispatch a) -> CWString -> CWString -> IO ()
foreign import ccall   "ReadProperty"           cReadProperty      :: (IDispatch a) -> CWString -> IO CWString
foreign import ccall   "Method_S_S"             cMethod_S_S        :: (IDispatch a) -> CWString -> CWString -> IO ()
foreign import ccall   "Method_S"               cMethod_S          :: (IDispatch a) -> CWString -> IO ()
foreign import ccall   "ReleaseObject"          cReleaseObject     :: (IDispatch a) -> IO ()
foreign import ccall   "stdlib.h free"          cfree              :: CWString -> IO ()
foreign import ccall   "stdlib.h free"          cDispatchFree      :: (IDispatch a) -> IO ()
foreign import stdcall "windows.h SysFreeString"  cSysFreeString   :: CWString -> IO ()
foreign import stdcall "ole2.h OleInitialize"     cOleInitialize   :: CInt -> IO ()
foreign import stdcall "ole2.h OleUninitialize"   cOleUninitialize :: IO ()

