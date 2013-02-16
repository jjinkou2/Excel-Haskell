{-# LANGUAGE ForeignFunctionInterface #-}
{- module Memory (
    Word32
    , allocBytes
    , coAlloc
    )
where

-}
import Data.Word ( Word16, Word32 )
import Data.Int ( Int32 )
import Foreign.Ptr    (nullPtr,castPtr,Ptr(..))
import HDirect (allocBytes)
import PointerPrim (primAllocMemory)
import Base




-- | @coAlloc sz@ allocates @sz@ bytes from the COM task allocator, returning a pointer.
-- The onus is on the caller to constrain the type of that pointer to capture what the
-- allocated memory points to.
coAlloc :: Word32 -> IO (Ptr a)
coAlloc sz = allocMemory sz

main = stringFromHR 13

--Primitives/helpers:

allocMemory :: Word32 -> IO (Ptr a)
allocMemory sz = do
  a <- primAllocMemory sz
  if a == nullPtr then
     ioError (userError "allocMemory: not enough memory")
   else
     return (castPtr a)
