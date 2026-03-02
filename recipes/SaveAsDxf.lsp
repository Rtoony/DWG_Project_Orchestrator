;; -----------------------------------------------------------------------------
;; SaveAsDxf.lsp
;;
;; Description:
;; This routine saves the currently active drawing file as a DXF file.
;; The new DXF will have the same base name and be saved in the same
;; directory as the original DWG file.
;;
;; Usage:
;; 1. Load this file into AutoCAD using the APPLOAD command or by dragging
;;    and dropping it into the drawing window.
;; 2. Type the command "SAVEASDXF" in the command line and press Enter.
;; -----------------------------------------------------------------------------
(defun C:SaveAsDxf ( / dwg_path dwg_name dxf_name)
  (princ "\nInitializing Save As DXF command...")

  ;; Get the directory path of the current drawing (e.g., "C:\\My Drawings\\")
  (setq dwg_path (getvar "dwgprefix"))

  ;; Get the base filename of the drawing without the .dwg extension
  (setq dwg_name (vl-filename-base (getvar "dwgname")))

  ;; Check if the drawing has been saved before. If not, dwg_path will be empty.
  (if (or (= dwg_path "") (= dwg_path nil))
    (progn
      (alert "Error: The drawing must be saved first before it can be exported to DXF.")
      (princ "\nCommand aborted. Please save the drawing.")
    )
    (progn
      ;; Construct the full path and filename for the new DXF file
      ;; by concatenating the path, base name, and ".dxf" extension.
      (setq dxf_name (strcat dwg_path dwg_name ".dxf"))

      ;; Provide feedback to the user in the command line
      (princ (strcat "\nSaving file as: " dxf_name))

      ;; Execute the DXFOUT command non-interactively.
      ;; "._dxfout" ensures the command works in any language version of AutoCAD.
      ;; The first argument is the full filename for the output.
      ;; The second argument, "", accepts the default version/precision settings.
      ;; For specific versions, you can use (command "._dxfout" dxf_name "V" "R2018" "").
      (command "._dxfout" dxf_name "")

      (princ "\nDXF file created successfully.")
    )
  )
  
  ;; The (princ) function at the end suppresses the return value of the last
  ;; evaluated expression (e.g., "nil") from being printed to the command line,
  ;; resulting in a cleaner exit.
  (princ)
)

;; A message to confirm that the file has been loaded successfully.
(princ "\n'SaveAsDxf.lsp' loaded. Type SAVEASDXF to run.")
(princ)