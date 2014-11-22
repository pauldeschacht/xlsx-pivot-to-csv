(ns xlsx-pivot-to-csv.core-test
  (:require [clojure.test :refer :all]
            [xlsx-pivot-to-csv.core :refer :all]))

(deftest extract-pivot-cache
  (testing "extract-pivot-cache-from-xlsx-file"
    (let [xlsx-file "test/pivot_table_example.xlsx"
          defs-file "xl/pivotCache/pivotCacheDefinition1.xml"
          recs-file "xl/pivotCache/pivotCacheRecords1.xml"
          data (extract-pivot-data xlsx-file defs-file recs-file)]
      (are [x y] (= x y) 
            (vec (nth data 0)) [{:type :n, :column 0, :value "1012"} {:type :x, :column 1, :value "REPUBLICAN"} {:type :n, :column 2, :value "2408"} {:type :x, :column 3, :value "71 +"} {:type :s, :column 4, :value "08/2006"} {:type :x, :column 5, :value "51"} {:type :s, :column 6, :value "PERM"}]
            (vec (nth data 1)) [{:type :n, :column 0, :value "1013"} {:type :x, :column 1, :value "REPUBLICAN"} {:type :n, :column 2, :value "2411"} {:type :x, :column 3, :value "71 +"} {:type :s, :column 4, :value "08/2006"} {:type :x, :column 5, :value "50"} {:type :s, :column 6, :value "PERM"}]
            (vec (nth data 2)) [{:type :n, :column 0, :value "1014"} {:type :x, :column 1, :value "DEMOCRAT"} {:type :n, :column 2, :value "2424"} {:type :x, :column 3, :value "71 +"} {:type :s, :column 4, :value "08/2006"} {:type :x, :column 5, :value "50"} {:type :s, :column 6, :value "PERM"}]
            (vec (nth data 3)) [{:type :n, :column 0, :value "1015"} {:type :x, :column 1, :value "DEMOCRAT"} {:type :n, :column 2, :value "2418"} {:type :x, :column 3, :value "71 +"} {:type :s, :column 4, :value "08/2006"} {:type :x, :column 5, :value "50"} {:type :s, :column 6, :value "POLL"}]
           (vec (nth data 4)) [{:type :n, :column 0, :value "1016"} {:type :x, :column 1, :value "REPUBLICAN"} {:type :n, :column 2, :value "2411"} {:type :x, :column 3, :value "71 +"} {:type :s, :column 4, :value "08/2006"} {:type :x, :column 5, :value "50"} {:type :s, :column 6, :value "PERM"}]
            
           
))))
