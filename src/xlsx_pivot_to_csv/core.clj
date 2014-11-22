(ns xlsx-pivot-to-csv.core
  (:use [clojure.tools.cli :refer [parse-opts]])
  (:require [clojure.java.io :as io])
  (:require [clojure.data.xml :as xml])
  (:require [clojure.zip :as zip])
  (:require [clojure.data.zip.xml :as zip-xml])
  (:gen-class)
  )

;;
;; PIVOTCACHEDEFINITIONS
;;
(defn get-values [cachefield-node]
  "extract the values from a single cachefield element"
  (->> cachefield-node
       (zip/xml-zip)
       (zip/down) ;; sharedItems
       (zip/children) ;; lazy-seq of <sharedItems>..</shareItems>
       (map #(get-in % [:attrs :v]))))

(defn transform-column-definitions [cachefields]
  "transform each <cachefield> into a map with the entries
:name name of the column
:values optional values of the column
:column index of the column"
  (map (fn [cachefield index]
         {:name (get-in cachefield [:attrs :name] )
          :values (get-values cachefield)
          :column index
          })
       cachefields
       (range)))

(defn extract-cachefields [inputstream]
  "read the xml formatted input stream and extract all the <cachefield> entries"
  (-> inputstream
      (xml/parse)
      (zip/xml-zip)
      (zip-xml/xml-> :cacheFields
                     :cacheField)))

(defn extract-defs [inputstream]
  "read the xml formatted file with the pivotcachedefinitions and return a map for each definition"
  (->> inputstream
       (extract-cachefields)
       (map first)
       (transform-column-definitions)))
;;
;; ROWS
;;
;; row is a list of map {:type :column :value}
;; type is the xml tag x for reference, s for string, n for numeric
;; column is the column index
;; value is the value of the cell (defined as string, no conversion is done)
(defn extract-row-values [node]
  (map (fn [node index]
         {:type (:tag node)
          :column index
          :value (get-in node [:attrs :v])
                   })
       (zip/children (zip/xml-zip node))
       (range))
  )

(defn extract-rows [inputstream]
  (->> inputstream
       (xml/parse)
       (zip/xml-zip)
       (zip/children) ;;lazy-seq of <r>..</r>
       (map extract-row-values)))


(defn join-ref-data [definitions row]
  "Each cell with type :x contains a reference to a value in the corresponding column definition. The actual value is retrieved from the column defintion and merged with the cell"
  (map (fn [cell]
         (if (= (:type cell) :x)
           (let [definition (nth definitions (:column cell))
                 values (:values definition)]
             (merge cell {:value (nth values (Integer. (:value cell)))}))
           cell
           ))
       row))

(defn join-data-with-definitions [definitions rows]
  (map #(join-ref-data definitions %) rows))

(defn pivot-data-to-csv [rows sep]
  "convert the data into a csv format"
  (map (fn [row]
         (doall (apply str
                       (interpose sep (map :value row)))))
       rows))

(defn process-zip-inputstream [xlsx-filename filename f]
  "the xlsx file is nothing more than a zip file. Unzip the file, extract the required file and process the content with the function f"
  (with-open [zipfile (java.util.zip.ZipFile. xlsx-filename)]
    (let [zipentry (.getEntry zipfile filename)
          inputstream (.getInputStream zipfile zipentry)]
      (doall (f inputstream)) ;; force evaluation before the file is closed
      )))

(defn extract-pivot-data [xlsx-file def-file rec-file]
  (let [defs (process-zip-inputstream xlsx-file def-file extract-defs)
        rows (process-zip-inputstream xlsx-file rec-file extract-rows)]
    (join-data-with-definitions defs rows)
    )
  )

(defn write-data [lines file]
  (with-open [out (io/writer file :append false)]
    (doall
     (doseq [line lines]
       (.write out line)
       (.write out "\n")))))

(defn default-output-name [in]
  (clojure.string/replace in #"(?i)xlsx$" "csv")
  )

(def cli-options
  [
   ["-x" "--xlsx XLSX" "input XLSX file"]
   ["-o" "--out OUT" "output file (csv)"]
   ["-d" "--defs DEFS" "definition filename, defaults to xl/pivotCache/pivotCacheDefinition1.xml" :default "xl/pivotCache/pivotCacheDefinition1.xml"]
   ["-r" "--recs ROWS" "records filename, defaults to xl/pivotCache/pivotCacheRecords1.xml" :default "xl/pivotCache/pivotCacheRecords1.xml"]
   ])

(defn -main [& args]
  (let [options (parse-opts args cli-options)
        xlsx-file (:xlsx (:options options))
        defs-file (:defs (:options options))
        recs-file (:recs (:options options))
        out (:out (:options options))
        out* (if (nil? out) (default-output-name xlsx-file) out)
        ]
    (-> (extract-pivot-data xlsx-file defs-file recs-file)
        (pivot-data-to-csv "^")
        (write-data out*))))


  
