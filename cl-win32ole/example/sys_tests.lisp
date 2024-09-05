(in-package :cl-win32ole-sys)

(let ((prog-id "BricscadApp.AcadApplication"))
(cffi:with-foreign-objects ((clsid 'CLSID))
  (with-ole-str (s prog-id)
    (progn 
    (format t "~a~%" (clsid-from-prog-id s clsid)))
  )))



(clsid-from-prog-id "BricscadApp.AcadApplication")
