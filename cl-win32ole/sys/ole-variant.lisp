;;;; -*- Mode: LISP; Syntax: COMMON-LISP; -*-
(in-package :cl-win32ole-sys)

(defconstant VT_EMPTY	 0)
(defconstant VT_NULL	 1)
(defconstant VT_I2	 2)
(defconstant VT_I4	 3)
(defconstant VT_R4	 4)
(defconstant VT_R8	 5)
(defconstant VT_CY	 6)
(defconstant VT_DATE	 7)
(defconstant VT_BSTR	 8)
(defconstant VT_DISPATCH 9)
(defconstant VT_ERROR	 10)
(defconstant VT_BOOL	 11)
(defconstant VT_VARIANT	 12)
(defconstant VT_UNKNOWN	 13)
(defconstant VT_DECIMAL	 14)
(defconstant VT_I1	 16)
(defconstant VT_UI1	 17)
(defconstant VT_UI2	 18)
(defconstant VT_UI4	 19)
(defconstant VT_I8	 20)
(defconstant VT_UI8	 21)
(defconstant VT_INT	 22)
(defconstant VT_UINT	 23)
(defconstant VT_VOID	 24)
(defconstant VT_HRESULT	 25)
(defconstant VT_PTR	 26)
(defconstant VT_SAFEARRAY 27)
(defconstant VT_CARRAY	 28)
(defconstant VT_USERDEFINED 29)
(defconstant VT_LPSTR	 30)
(defconstant VT_LPWSTR	 31)
(defconstant VT_RECORD	 36)
(defconstant VT_INT_PTR	 37)
(defconstant VT_UINT_PTR 38)
(defconstant VT_FILETIME 64)
(defconstant VT_BLOB	 65)
(defconstant VT_STREAM	 66)
(defconstant VT_STORAGE	 67)
(defconstant VT_STREAMED_OBJECT	 68)
(defconstant VT_STORED_OBJECT	 69)
(defconstant VT_BLOB_OBJECT	 70)
(defconstant VT_CF	 71)
(defconstant VT_CLSID	 72)
(defconstant VT_VERSIONED_STREAM 73)
(defconstant VT_BSTR_BLOB #xfff)
(defconstant VT_VECTOR	 #x1000)
(defconstant VT_ARRAY	 #x2000)
(defconstant VT_BYREF	 #x4000)
(defconstant VT_RESERVED #x8000)
(defconstant VT_ILLEGAL	 #xffff)
(defconstant VT_ILLEGALMASKED #xfff)
(defconstant VT_TYPEMASK #xfff)

(defconstant VARIANT_TRUE  -1)
(defconstant VARIANT_FALSE  0)

(cffi:defctype VARTYPE :unsigned-short)
(cffi:defctype VARIANT_BOOL :short)
(cffi:defctype SCODE LONG)

(cffi:defcunion variant-union
  (long-long LONGLONG)
  (long :long)
  (bool VARIANT_BOOL)
  (int :int)
  (float :float)
  (double :double)
  (date :double)
  (pdate :pointer)
  (pointer :pointer))

(cffi:defcstruct variant
  (vt VARTYPE)
  (wReserved1 WORD)
  (wReserved2 WORD)
  (wReserved3 WORD)
  (value variant-union))

(cffi:defctype VARIANTARG VARIANT)

(defun variant-type* (variant)
  (cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'vt))

(defun variant-type (variant)
  (logand (variant-type* variant) (lognot VT_BYREF)))

(defun (setf variant-type) (new-type variant)
  (setf (cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'vt) new-type))

(defun variant-array-p (variant)
  (not (zerop (logand (variant-type* variant) VT_ARRAY))))

(defun variant-byref-p (variant)
  (not (zerop (logand (variant-type* variant) VT_BYREF))))

(defun variant-union-accessor (variant)
  (cond ((variant-array-p variant)
         'pointer)
        (t (eswitch (variant-type variant)
             (VT_EMPTY 'long-long)
             (VT_I4 'long)
             (VT_R4 'float)
             (VT_R8 'double)
             (VT_DATE (if (variant-byref-p variant) 'pdate 'date))
             (VT_BSTR 'pointer)
             (VT_BOOL 'bool)
             (VT_VARIANT 'pointer)
             (VT_DISPATCH 'pointer)
             ))))

(defun variant-value (variant)
  (cffi:foreign-slot-value
   (cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'value)
   `(:union ,'variant-union)
   (variant-union-accessor variant)))

(defun variant-array-value (variant)
  (let ((pointer (variant-value variant)))
    (if (variant-byref-p variant)
        (cffi:mem-aref pointer :pointer)
        pointer)))

(defun (setf variant-value) (new-value variant)
  (setf
   (cffi:foreign-slot-value
    (cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'value)
    `(:union ,'variant-union)
    (variant-union-accessor variant))
   new-value))

(defun alloc-variant ()
  (let ((variant (cffi:foreign-alloc 'VARIANT)))
    (dformat t "variant::alloc ~a~%" variant)
    (VariantInit variant)
    variant))

(defun free-variant (variant)
  (dformat t "variant::VariantClear ~a~%" variant)
  (succeeded (VariantClear variant))
  (dformat t "variant::free ~a~%" variant)
  (cffi-sys:foreign-free variant))

(defun variant-copy (dest src)
  (succeeded (VariantCopy dest src))
  dest)

(defmacro with-ole-str ((var str) &body body)
  "wchar_t string which ends of null"
  (let ((len (gensym)))
    `(let* ((,len (length ,str)))
       (cffi:with-foreign-object (,var :unsigned-short (1+ ,len))
         (loop for i from 0 below ,len
            do (setf (cffi:mem-aref ,var :unsigned-short i)
                     (char-code (char ,str i)))
            finally (setf (cffi:mem-aref ,var :unsigned-short ,len) 0))
         ,@body))))

(defun bstr->lisp (bstr)
  (with-output-to-string (out)
    (loop for i from 0
       for *ptr = (cffi:mem-aref bstr 'WORD i)
       while (not (zerop *ptr))
       do (write-char (code-char *ptr) out))))

(defun lisp->bstr (lisp)
  (with-ole-str (ole-str lisp)
    (SysAllocString ole-str)))

(defun calculate-variant-type (value)
  (cond ((or (eq nil value) (eq t value))
         VT_BOOL)
        ((cffi-sys:pointerp value)
         VT_DISPATCH)
        (t (etypecase value
             (string VT_BSTR)
             (fixnum VT_I4)
             (single-float VT_R4)
             (double-float VT_R8)))))

(defun lisp->variant (value &optional type)
  (unless type (setf type (calculate-variant-type value)))
  (let ((variant (alloc-variant)))
    (setf (variant-type variant) type)
    (setf (variant-value variant)
          (eswitch (variant-type variant)
            (VT_I4 value)
            (VT_R4 value)
            (VT_R8 value)
            (VT_BSTR (lisp->bstr value))
            (VT_BOOL (if value VARIANT_TRUE VARIANT_FALSE))
            (VT_VARIANT value)
            (VT_DISPATCH value)
            ((logior VT_VARIANT VT_ARRAY) value)
            ))
    variant))
