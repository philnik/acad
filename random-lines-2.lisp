
(ql:quickload "trivial-dump-core")

(eval-when (:compile-toplevel :load-toplevel :execute)
  (asdf:oos 'asdf:load-op :cl-win32ole)
  (use-package :cl-win32ole))


(defun hello ()

(let ((acad (create-object "BricscadApp.AcadApplication")))
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

)



(trivial-dump-core:save-executable "./random-lines.exe" #'hello)
