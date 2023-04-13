(push "c:/users/me/Application Data/lisp/cl-win32ole/" asdf:*central-registry*)

(ql:quickload 'cl-win32ole)

(use-package 'cl-win32ole)
(use-package 'cl-win32ole-sys)

(in-package :cl-win32ole)


(format t "~a" (random-point))

(defun random-point ()
  (with-output-to-string (str)
    (format str "~d,~d,~d" (random 10000) (random 1000)  (random 1000) )))


(defun m ()
  (loop for i from 1 to 100
	do (let ((spoint (random-point))
		 (epoint (random-point)))
	     (invoke model :Addline spoint epoint)
	     ))
  )



(defun m ()
  (loop for i from 1 to 100
	do (let ((spoint (random-point))
		 (epoint (random-point)))
	     (invoke model :Addline spoint epoint)
	     ))
  )



(defun k (j)
  (loop for i from 1 to j
	do (let ((spoint (random-point))
		 (epoint (random-point)))
	     (invoke model :Addline spoint epoint)
	     ))
  )



(m)


(progn
  (cl-win32ole-sys::co-initialize (cffi-sys:null-pointer))
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
  (m)
  )

(m)

(defun zoomall ()
   (invoke (ole doc :Application) :ZoomAll)
  )
