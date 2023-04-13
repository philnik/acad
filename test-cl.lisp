





(push "c:/users/me/Application Data/lisp/cl-win32ole/" asdf:*central-registry*)

(ql:quickload 'cl-win32ole)

(use-package 'cl-win32ole)
(use-package 'cl-win32ole-sys)

(defun zoomall ()
(progn
  (cl-win32ole-sys::co-initialize (cffi-sys:null-pointer))
  (let ((acad (create-object "BricscadApp.AcadApplication")))
    (progn
      (format t "acad:~a~%" acad)
      (format t "acad.ActiveDocument:~a~%" (ole acad :ActiveDocument))
      (format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))
      (setf doc (ole acad :ActiveDocument))
      (format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))
      (format t "acad.ActiveDocument.ZoomAll():~a~%" "1")
      (invoke (ole doc :Application) :ZoomAll)
      (format t "doc.Name:~a~%" (ole doc :Name))
      ))
  (cl-win32ole-sys::co-uninitialize)
  )
)
