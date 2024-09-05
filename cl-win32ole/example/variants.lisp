
(in-package :cl-user)


;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))


(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)

(in-package :cl-win32ole-sys)
  

(setf a (cffi:foreign-alloc :int :initial-contents '(1 2 3)))
(loop for i from 0 below 3
      collect (cffi:mem-aref a :int i))

(setf d2 (make-array 2 :element-type 'double-float :initial-contents #(3.0d0 4.0d0)))
(setf a (cffi:foreign-alloc :double :initial-contents d2))
(loop for i from 0 below 2
      collect (cffi:mem-aref a :double i))



(setf d2 (make-array 2 :element-type 'double-float :initial-contents #(3.0d0 4.0d0)))
(setf v0 (alloc-variant))
(setf (variant-type v0) (logior VT_ARRAY VT_R8))
(setf p1 (cffi:foreign-alloc :double :initial-contents d2))
(setf (variant-value v0) p1)

(cffi:mem-aref p1 :double 0)
(cffi:mem-aref p1 :double 1)
;;;(3,4)


(cffi:mem-aref (variant-value v0) :double 0)
(cffi:mem-aref (variant-value v0) :double 1)


(setf (cffi:mem-aref (variant-value v0) :double 0) 15.0d0)
(setf (cffi:mem-aref (variant-value v0) :double 1) 50.0d0)
(cffi:mem-aref (variant-value v0) :double 0) 
(cffi:mem-aref (variant-value v0) :double 1) 
;;;15,50


