


(ql:quickload "trivial-dump-core")
;(declaim (sb-ext:muffle-conditions cl:warning))
;(declaim (sb-ext:muffle-conditions cl:style-warning))
(push #'(lambda () (print (sb-kernel::dynamic-usage))) sb-kernel::*after-gc-hooks*)
(sb-kernel::gc)

;(ql:quickload "bt-semaphore")

;;(push "c:/users/me/Application Data/lisp/cl-win32ole/" asdf:*central-registry*)

;;(ql:quickload 'cl-win32ole)

(use-package 'cl-win32ole)

(use-package 'cl-win32ole-sys)

;(in-package :cl-win32ole)


(defun random-point (i j k)
  (with-output-to-string (str)
    (format str "~d,~d,~d" (random i) (random j)  (random k) )))

;(cl-win32ole-sys::co-initialize-ex (cffi-sys:null-pointer)  cl-win32ole-sys::COINIT_MULTITHREADED)
;(cl-win32ole-sys::co-initialize-ex (cffi-sys:null-pointer)  cl-win32ole-sys::COINIT_MULTITHREADED)

(defun hello ()
  (sb-thread:make-thread
   (progn
  (co-initialize-ex (cffi-sys:null-pointer) COINIT_MULTITHREADED)
  (setf acad (create-object "BricscadApp.AcadApplication"))
  (setf (ole acad :Visible) 1)
  (setf n (ole acad :ActiveDocument :Name))
  (format t "~a" n)
  (format t "acad:~a~%" acad)
  (format t "acad.ActiveDocument:~a~%" (ole acad :ActiveDocument))
  (format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))
  (setf doc (ole acad :ActiveDocument))
  (format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))
  (format t "acad.ActiveDocument.ZoomAll():~a~%" "1")
  (invoke (ole doc :Application) :ZoomAll)
  (format t "doc.Name:~a~%" (ole doc :Name))
  (setf doc (ole acad :ActiveDocument))
  (setf model (ole doc :ModelSpace))
  )
))

;(hello)

;(trivial-dump-core:save-executable "./random-lines.exe" #'hello)



(defun random-point (i j k)
  (with-output-to-string (str)
    (format str "~d,~d,~d" (random i) (random j)  (random k) )))


;CL-WIN32OLE-SYS:VARIANT
;(defparameter SLYNK:*COMMUNICATION-STYLE* nil)
;(defparameter SLYNK-BACKEND::PREFERRED-COMMUNICATION-STYLE nil)


(defun m (i j k)
;   (cl-win32ole-sys::co-initialize (cffi-sys:null-pointer))
  (loop for i from 1 to 10
	do (let ((spoint (random-point i j k))
		 (epoint (random-point (* 2 i) (* j 5) k)))
	     (invoke model :Addline spoint epoint)
	     ))
 ; (cl-win32ole-sys::co-uninitialize)
  )

(cl-win32ole-sys::co-initialize (cffi-sys:null-pointer))

(defun times (j)
  (dotimes (i j)
    (m (+ 1 (random 1000)) (+ 1 (random 10)) 100)
    )
  )

;(times 10)

(defun zoom ()
  (invoke (ole doc :Application) :ZoomAll)
)
  
(defun nnn (i j k)
  (progn 
     (m i j k)
     (zoom)
     ))


(defun get-doc-name (doc)
  (ole doc :Name))

(get-doc-name doc)

(defun point ( x y z)
  (with-output-to-string (str)
    (format str "~d,~d,~d" x y z))
  )


(defun make-box (w h)
  (let ((p0 (point 0 0 0))
	(p1 (point w 0 0))
	(p2 (point w h 0))
	(p3 (point 0 h 0))
	)
    (invoke model :Addline p0 p1)
    (invoke model :Addline p1 p2)
    (invoke model :Addline p2 p3)
    (invoke model :Addline p3 p0)
    ))


(defun  make-box-2 (w h)
  (let ((p0 (point (- 0 w) (- 0 h) 0))
	(p1 (point (- 0 w) (+ 0 h) 0))
	(p2 (point (+ 0 w) (+ 0 h) 0))
	(p3 (point (+ 0 w) (- 0 h) 0))
	)
    (invoke model :Addline p0 p1)
    (invoke model :Addline p1 p2)
    (invoke model :Addline p2 p3)
    (invoke model :Addline p3 p0)
    ))

;(make-box-2 300 300)

;(point 1 2 3)


					;(hello)
(defun make-lot-of-boxes ()
  (loop for i from 100 to 200
	do (make-box i i)
	)
  
  (zoom)
  )

(defun make-lot-of-boxes-2 ()
  (sb-thread:make-thread
  (loop for i from 100 to 1000
	do (make-box-2 i i)
	)
  )
  (zoom)
  )

;(make-lot-of-boxes-2)


					;(hello)
(defun make-lot-of-boxes-2 ()
  (mapcar #'(lambda (i) (make-box-2 i i)) '(100 200 250 260 270 300))
  )

;(make-lot-of-boxes-2)




(progn 
(cl-win32ole-sys::coinitialize)
(setf acad2 (create-instance "BricscadApp.AcadApplication"))
)
acad2
(setf n1 (ole acad2 :ActiveDocument :Name)



;(trivial-dump-core:save-executable "./random-lines2.exe" #'hello)
