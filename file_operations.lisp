
(in-package :cl-win32ole)

(defun send_command (str)
  (invoke doc :SendCommand str)
  )

(defun open_dwg (fname)
  (invoke (ole (ole doc :Application) :Documents) :Open fname)
  )

(setq fname2 "C:\\Users\\filip\\AppData\\Roaming\\blender\\agiou-nikolaou\\ag.nikolaou.dwg")
      
(open_dwg fname2)

(defun draw_concentric_circles (max)      
  (loop for i from 1 to max
	collect (let ((str (format nil "circle 0,0 ~d " i)))
		  (send_command str))))


  ;;(invoke doc :AddTable "0.0, 0.0" 10.0 10.0 10.0 100)

;(invoke model :AddMtext "200.0 200.0" 1000d0 "1hellodddd11")
;(invoke model :AddLine "0.0 0.0" "1.0 0.0")



(defun list-to-pointer (lst)
  (let* ((size (length lst))
         (ptr (cffi:foreign-alloc :double :count size)))
    (loop for i from 0 below size
          for elem in lst
          do (setf (cffi:mem-aref ptr :double i) elem))
    ptr))  

(
(setq a (list-to-pointer (list 0.0d0 0.0d0 0.0d0)))
  
 (eval `(invoke model :AddPoint ,a) )

 


(in-package :cl-win32ole-sys)
(cffi:defcfun ("process_array" process-array)
    ((arr :pointer) (length :int))
    (:void))
