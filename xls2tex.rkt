#lang racket

;; Copyright (c) 2017, Florian Biermann (fbie@itu.dk)

;; Permission is hereby granted, free of charge, to any person obtaining a copy
;; of this software and associated documentation files (the "Software"), to deal
;; in the Software without restriction, including without limitation the rights
;; to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
;; copies of the Software, and to permit persons to whom the Software is
;; furnished to do so, subject to the following conditions:

;; The above copyright notice and this permission notice shall be included in all
;; copies or substantial portions of the Software.

;; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
;; IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
;; FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
;; AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
;; LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
;; OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
;; SOFTWARE.

(require racket/cmdline)
(require racket/format)
(require xml)

;;; Construction of a usable row- and cell-representation.

(struct pos (row col)
  #:guard (λ (row col name)
            (unless (and (number? row) (number? col))
              (error "Not a valid pos."))
            (values row col))
  #:transparent)

(define zero (pos 0 0))

;; Advance column by one or set to n.
(define (xls/pos-next-col curr [n #f])
  (if n
      (pos (pos-row curr) (string->number n))
      (pos (pos-row curr) (add1 (pos-col curr)))))

;; Advance row by one or set to n.
(define (xls/pos-next-row curr [n #f])
  (if n
      (pos (string->number n) (pos-col curr))
      (pos (add1 (pos-row curr)) (pos-col curr))))

;; Filter elements recursively from a list xs using the given predicate.
(define (xls/filter-recursive predicate lst)
  (define (filter-recurse predicate lst)
    (if (list? lst)
        (filter predicate (map (curry xls/filter-recursive predicate) lst))
        lst))
  ;; Wrap predicate and include a check for empty lists, we don't want
  ;; to keep those.
  (filter-recurse (λ (x) (and (not (empty? x)) (predicate x))) lst))

;; Unwrap singleton lists to their single element. Does nothing to
;; empty lists.
(define (xls/unwrap lst)
  (if (= (length lst) 1)
      (car lst)
      lst))

;; Unwrap lists recursively.
(define (xls/flatten lst)
  (if (list? lst)
      (xls/unwrap (map xls/flatten lst))
      lst))

;; Read the XLS formatted file into an xexpr.
;; TODO: Handle non UTF-8 smartly.
;; TODO: Remove newline returns.
(define (xls/read filename)
  (let [(predicate (λ (x)
                     (not (and (string? x)
                               (regexp-match #px"\r\n\\s*" x)))))
        (doc (xml->xexpr (document-element (with-input-from-file filename read-xml))))]
    (xls/flatten (xls/filter-recursive predicate doc))))

;; A variant of assoc that is more robust to non-pairs in lists.
(define (xls/xexpr-assoc key dict)
  (if (or (not (list? dict)) (empty? dict))
      #f
      (let [(candidate (car dict))]
        (if (and (list? candidate) (eq? (car candidate) key))
            candidate
            (xls/xexpr-assoc key (cdr dict))))))

;; Get value for some key from an xls structure. Returns #f if the key
;; doesn't exist.
(define (xls/get key dict)
  (let [(val (xls/xexpr-assoc key dict))]
    (if val
        (cadr val)
        #f)))

;; Retrieve a formula from an element, if any.
(define (xls/get-formula elem)
  (xls/get 'ss:Formula (car elem)))

;; Retrieve an index offset from an element.
(define (xls/get-index elem)
  (or (xls/get 'ss:Index elem) (xls/get 'ss:Index (car elem))))

;; Retrieve data from an element, if any.
(define (xls/get-data elem)
  (caddr (xls/xexpr-assoc 'Data elem)))

;; Retrieve the element's name tag.
(define (xls/get-name elem)
  (xls/get 'ss:Name elem))

;; A cell consists of an expression and a position.
;; TODO: Maybe only use column?
(struct cell (expr pos)
  #:guard (λ (expr pos name)
            (unless (and (string? expr) (pos? pos))
                (error "Not a valid cell."))
              (values expr pos))
  #:transparent)

;; Return a string that represents a cell in the spreadsheet.
(define (xls/parse-cell c p)
  (let* [(ccell (cdr c))
         (formula (xls/get-formula ccell))
         (index (xls/get-index ccell))
         (data (xls/get-data ccell))]
    (cell (or formula data) (xls/pos-next-col p index))))

;; Retrieve all elements from the xls that are of the given type.
(define (xls/filter-type type lst)
  (filter (λ (elem) (and (list? elem) (eq? (car elem) type))) lst))

;; Typical prefix-sum. If lst has n elements, then the result has n +
;; 1, where the first element is the original prefix.
(define (scan f pre lst)
  (match (length lst)
    [0 (list pre)]
    [_ (cons pre (scan f (f pre (car lst)) (cdr lst)))]))

;; Return a string that represents a row in TeX table format together
;; with the last position.
(define (xls/row->cells row pos)
  (unless (list? row)
    (error "Row must be a list!"))
  (cdr (scan (λ (p c)
               (unless (cell? p)
                 (error "Prefix not a cell."))
               (xls/parse-cell c (cell-pos p)))
             (cell "" pos) ;; Fake cell that we strip later.
             row)))

;; Get the first sheet with the correct sheet name. There should not
;; be two sheets with the same name.
(define (xls/get-sheet sheetname doc)
  (car (filter (λ (sheet)
                 (string=? sheetname (xls/get-name sheet)))
               (xls/filter-type 'Worksheet doc))))

(define (xls/sheet->cells sheet doc)
  (let* [(rows (xls/filter-type 'Row (xls/xexpr-assoc 'Table (xls/get-sheet sheet doc))))
         ;; Compute row numbers; allows us to use map instead of scan.
         (irows (cdr (scan (λ (pre row) (cons (xls/pos-next-row (car pre) (xls/get-index row)) row))
                           (cons zero empty)
                           rows)))]
    (for/list ([irow irows])
      (let [(pos (car irow))
            (row (cdr irow))]
        (xls/row->cells (xls/filter-type 'Cell row) pos)))))


;;; Cell reference formatting.

;; Works up to column 26. I am too lazy to implement a general
;; solution, and how likely is it that our spreadsheets exceed 26
;; columns?
;; TODO: Make work for columns > 26.
(define (c1->column col)
    (string (integer->char (+ 64 col))))

(struct cellref (col row cabs rabs)
  #:guard (λ (col row cabs rabs name)
            (unless (and (number? col) (number? row) (boolean? cabs) (boolean? rabs))
              (error "Not a valid cellref."))
            (values col row cabs rabs))
  #:transparent)

(define (cellref-row-abs? ref)
  (cellref-rabs ref))

(define (cellref-col-abs? ref)
  (cellref-cabs ref))

;; Format for R1C1. This is easy, we only need to decide on the
;; formatting string.
(define (print-r1c1 ref)
  (let [(r (~a (format (if (cellref-row-abs? ref) "R~a" "R[~a]") (cellref-row ref))))
        (c (~a (format (if (cellref-col-abs? ref) "C~a" "C[~a]") (cellref-col ref))))]
    (string-join (list r c) "")))

;; Format for A1. This is more complicated, because we need to
;; compute relative references wrt. the current position.
(define (print-a1 ref pos)
  (let [(r (if (cellref-row-abs? ref)
               (~a "$" (cellref-row ref))
               (number->string (+ (cellref-row ref) (pos-row pos)))))
        (c (if (cellref-col-abs? ref)
               (~a "$" (c1->column (cellref-col ref)))
               (c1->column (+ (cellref-col ref) (pos-col pos)))))]
    (string-join (list c r) "")))

;; Print a cell reference using either 'R1C1 or 'A1 styles. If 'A1 is
;; desired, you need to pass in a valid position, too.
(define (cellref-print style ref [pos zero])
  (match style
    ['R1C1 (print-r1c1 ref)]
    ['A1 (print-a1 ref pos)]))

;; Parse a R1C1 formatted cell reference and return a r1c1ref struct.
(define (parse-r1c1 refstring)
  (let [(get-num (λ (s) (string->number (car (regexp-match #px"-?\\d+" s)))))
        (row-abs (regexp-match #px"R\\d+" refstring))
        (row-rel (regexp-match #px"R\\[-?\\d+\\]" refstring))
        (col-abs (regexp-match #px"C\\d+" refstring))
        (col-rel (regexp-match #px"C\\[-?\\d+\\]" refstring))]
    (cellref (if (or col-rel col-abs)
                 (get-num (car (or col-rel col-abs)))
                 0)
             (if (or row-rel row-abs)
                 (get-num (car (or row-rel row-abs)))
                 0)
             col-abs
             row-abs)))

;; Convert a R1C1 formatted refstring to an A1 formatted refstring
;; using pos.
(define (r1c1->a1 refstring pos)
  (cellref-print 'A1 (parse-r1c1 refstring) pos))

;; Take all pairs from reps and apply them to str using
;; string-replace. The result is str with all occurrences of the first
;; element replaced with the second element for all pairs in reps.
(define (replace-all reps str)
  (if (and reps (list? reps) (not (empty? reps)))
      (replace-all (cdr reps) (string-replace str (caar reps) (cdar reps)))
      str))

;; Convert an R1C1 formula into an A1 formula for some absolute
;; position.
(define (formula/r1c1->a1 formula pos)
  (let* [(refs (regexp-match* #px"R\\[?-?\\d*\\]?C\\[?-?\\d*\\]?" formula))
         (reps (if refs
                   (map (λ (ref) (cons ref (r1c1->a1 ref pos))) refs)
                   empty))]
    (replace-all reps formula)))

;; Convert all cells in this collection of rows into A1 reference
;; format.
(define (sheet/r1c1->a1 rows)
  (for/list ([r rows])
    (for/list [(c r)]
      ;; Avoid unnecessary parsing and memory allocation.
      (if (string-prefix? (cell-expr c) "=")
          (cell (formula/r1c1->a1 (cell-expr c) (cell-pos c)) (cell-pos c))
          c))))



;;; Generating TeX code.

(define tex-newline " \\\\ \\hline")
(define tex-tabsep " &")
(define tex-cbar "c|")

(define (string-repeat n s)
  (string-append* (make-list n s)))

(define (tex/cell->tex c [cols 0])
  (~a (string-repeat (- (pos-col (cell-pos c)) cols) tex-tabsep)
      " \\texttt{"
      (cell-expr c)
      "}"))

(define (tex/beginline row)
  (~a row tex-tabsep))

(define (tex/tab-format cols)
  (~a "|" (string-repeat (add1 cols) tex-cbar)))

(define (tex/sheet-header cols)
  (~a tex-tabsep
      " "
      (string-join (build-list cols (λ (n) (c1->column (add1 n))))
                   (~a tex-tabsep " "))))

(define (xls/columns rows)
  (foldl max 0 (map pos-col (map cell-pos (map last rows)))))

(define (tex/print-sheet sheet)
  (let [(cols (xls/columns sheet))]
    (display (~a "\\begin{tabular}[" (tex/tab-format cols) "]"))
    (newline)
    (display (tex/sheet-header cols))
    (for/fold ([r 1])
              ([row sheet])
      ;; Newline comes always before cell contents.
      (display tex-newline)
      (newline)
      (display (tex/beginline r))
      (for/fold ([c 1])
                ([cell row])
        (display (tex/cell->tex cell c))
        c + 1)
      r + 1))
  (display tex-newline)
  (newline)
  (display "\\end{tabular}")
  (newline))



;;; Command line arguments and execution.

(define sheet-name (make-parameter "Sheet1"))
(define r1c1 (make-parameter #f))

(define xls-file
  (command-line
   #:program "xls2tex"
   #:once-each
   [("-s" "--sheet")
    sheet
    "Name of the sheet to convert to LaTeX code."
    (sheet-name sheet)]

   [("-r" "--r1c1")
    "Use R1C1 reference style instead of A1."
    (r1c1 #t)]

   #:args (filename)
   filename))

(let ([sheet (xls/sheet->cells (sheet-name) (xls/read xls-file))])
  (tex/print-sheet (if (r1c1)
                       sheet
                       (sheet/r1c1->a1 sheet))))
