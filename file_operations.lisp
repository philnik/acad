
(in-package :cl-win32ole)

(ql:quickload "cl-interpol")
(named-readtables:in-readtable :interpol-syntax)
;;;(let ((a 42))
;;;    #?"foo: \xC4\N{Latin capital letter U with diaeresis}\nbar: ${a}"

(defparameter +template+ "C:/Users/filip/AppData/Local/Bricsys/BricsCAD/V24x64/en_US/Templates/Default-mm.dwt")
(defparameter *ROOT* "C:/Users/filip/AppData/Roaming/blender/")
(setq fname2 "C:\\Users\\filip\\AppData\\Roaming\\blender\\agiou-nikolaou\\ag.nikolaou.dwg")
(setq fname3 "C:/Users/filip/AppData/Roaming/blender/empty3.dwg")


(defun send_command (str)
  (invoke doc :SendCommand str)
  )

(defun open_dwg (fname)
  (invoke (ole (ole doc :Application) :Documents) :Open fname)
  )

(invoke (ole doc :TextStyles) :Add "Standard1")

(defun save_dwg_as (fname)
  (send_command (format nil "(vla-saveas (vla-get-ActiveDocument (vlax-get-Acad-Object)) \"~a\" ) " fname3))
  )

(defun new_document (fname3)
(progn
  (open_dwg +template+)
  (send_command (format nil "(vla-saveas (vla-get-ActiveDocument (vlax-get-Acad-Object)) \"~a\" ) " fname3))
  ))

;;(new_document #?"${*ROOT*}/new_doc2.dwg")

(defun draw_concentric_circles (max)      
  (loop for i from 1 to max
	collect (let ((str (format nil "circle 0,0 ~d " i)))
		  (send_command str))))

;(draw_concentric_circles 10)
;(invoke model :AddMtext "200.0 200.0" 1000d0 "1hellodddd11")

;(invoke model :AddLine "0.0 0.0" "1.0 0.0")

(defun add_new_style_command ()
(send_command 
 #?"
(defun new_style ()
(entmake
(list
'(0 . \"STYLE\")
'(100 . \"AcDbSymbolTableRecord\")
'(100 . \"AcDbTextStyleTableRecord\")
'(2 . \"Design-Level2\")
'(70 . 0)
'(40 . 0.0)
'(41 . 1.0)
'(50 . 0.0)
'(71 . 0)
'(42 . 0.09375)
'(3 . \"ArialN.ttf\")
'(4 . \"\")
)
)) 
"
)
