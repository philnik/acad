(ql:quickload  "fiveam")
(push "c:/users/me/Application Data/lisp/cl-win32ole/" asdf:*central-registry*)
(eval-when (:compile-toplevel :load-toplevel :execute)
  (asdf:oos 'asdf:load-op :cl-win32ole-sys)
  (asdf:oos 'asdf:load-op :fiveam))

(in-package :cl-win32ole-sys)

(5am:def-suite cl-win32ole-sys-test)

(5am:in-suite cl-win32ole-sys-test)

(5am:test variant-test

  (5am:is (= 16 (cffi:foreign-type-size 'VARIANT)))

  (let ((v (alloc-variant)))
    (unwind-protect
         (progn
           (5am:is (= VT_EMPTY (variant-type v))
                   )
           (setf (variant-type v) VT_BOOL)
           (5am:is (= VT_BOOL (variant-type v)))
           (setf (variant-value v) VARIANT_TRUE)
           (5am:is (= VARIANT_TRUE (variant-value v)))
           (setf (variant-value v) VARIANT_FALSE)
           (5am:is (= VARIANT_FALSE (variant-value v)))

           (setf (variant-type v) VT_I4
                 (variant-value v) 123)
           (5am:is (= VT_I4 (variant-type v)))
           (5am:is (= 123 (variant-value v)))
           )
      (free-variant v)))
  )

(5am:test type-lib
 (let* ((dispatch (create-instance "WScript.Shell"))
        (type-info (dispatch-get-type-info dispatch))
        (type-lib (type-info-get-containing-type-lib type-info)))
   (unwind-protect
        (progn
          )
     (unknown-release type-lib)
;    (type-info-add-ref type-info)
     (dispatch-add-ref dispatch)
     )))


(5am:run! 'cl-win32ole-sys-test)
