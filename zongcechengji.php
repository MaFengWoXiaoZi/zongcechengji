<?php
    function oneExcelFileUploads($excelFile)
    {
        // 对上传的文件进行类型和大小检查, 判断其是否为excel文件类型, 是否超过1M大小, 判断文件是否已经存在
        if (
            ($_FILES[$excelFile]["type"] == "application/vnd.ms-excel"
        || $_FILES[$excelFile]["type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
         && $_FILES[$excelFile]["size"] <= 1024000)
        {
            if ($_FILES[$excelFile]["error"] > 0)
            {
                echo "错误: " . $_FILES[$excelFile]["error"] . "<br />";
            }
            else
            {
                if (file_exists(iconv("UTF-8","GBK//IGNORE", "D:\\AppServ\\www\\excel\\" . $_FILES[$excelFile]["name"])))
                {
                    echo $_FILES[$excelFile]["name"] . "已经存在";
                }
                else
                {
                    move_uploaded_file($_FILES[$excelFile]["tmp_name"],
                    iconv("UTF-8","GBK//IGNORE", "D:\\AppServ\\www\\excel\\" . $_FILES[$excelFile]["name"])); // windows下需要对文件名进行转码, 将utf-8编码转化为gbk, Linux默认为utf-8, 因此不需要
                    echo "Stored in: " . "excel\\" . $_FILES[$excelFile]["name"];
                }
            }
        }
        else
        {
            echo "<script>alert('您上传的不是excel文件或者excel文件大小超过了1M');</script>";
        }
    }

    oneExcelFileUploads("uploadedFile1");
    oneExcelFileUploads("uploadedFile2");
  
?>
<?php
    include "PHPExcel-1.8/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php";
    include "PHPExcel-1.8/PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php";
    date_default_timezone_set("PRC");

    echo "Stored in: " . $_FILES["uploadedFile1"]["tmp_name"];
    // 打开excel文件函数, 返回文件对象
    function openExcelFile($inputFileName)
    {
        try {
            $inputFileType = PHPExcel_IOFactory::identify($inputFileName); // 甄别文件类型
            $objReader = PHPExcel_IOFactory::createReader($inputFileType); // 跟据文件类型创建读取器
            $objReader->setReadDataOnly(true); // 只读取数据，忽略其它各种格式设置
            $objPHPExcel = $objReader->load($inputFileName); // 用读取器读取文件
        } catch(Exception $e) {
            die ("加载文件发生错误: " . pathinfo ($inputFileName1, PATHINFO_BASENAME) . ":" . $e->getMessage()); // 出错时，触发异常
        }
        return $objPHPExcel;
    }

    // 获得一个学生一学期的所有课程分数信息
    function getOneStudentGradesInfo($oneStudentSeq, $sheet)
    {
        $highestColumn = $sheet->getHighestColumn(); // 获得这张表的最大列数 
        $allColumn = PHPExcel_Cell::columnIndexFromString($highestColumn);  // 列数用英文字母表示, 当其大于Z时，需要转换成数字，否则会出错
        $oneStudentGradesInfo = array();  // 存放单个学生的所有课程分数信息
        $cellValue = NULL; // 判断获取的单元格值是否为空, 初始化为空
        $oneClassInfo = ""; // 一门课程的所有信息
        $oneClassAttr = 1; // 课程性质, 必修课为1, 选修课为2
        $oneClassCredit = 0; // 存放一门课程的学分
        $oneClassGrade = 0; // 存放分数信息的临时变量
        $index = 0; // 存放单个学生的所有课程信息的数组的下标
        for ($column = 4; $column <= $allColumn; $column++) {

            $columnA = PHPExcel_Cell::stringFromColumnIndex($column); // 列的数字下标需要转回字母下标
            $cellValue = $sheet->getCell($columnA . $oneStudentSeq)->getValue();

            // 获取当前分数单元格的值, 为空表示学生没有修该门课程, 因此不需要记录课程相关信息
            if ($cellValue == NULL)
            {
                continue;
            }else {
                $oneClassGrade = (int)$cellValue; // 记录一门课程的分数
                $oneClassInfo = $sheet->getCell($columnA . 3)->getValue(); // 记录一门课程的详细信息
                $oneClassCredit = (float)(end(explode("/", $oneClassInfo))); // 记录一门课程的学分
                if (preg_match("/必/", $oneClassInfo)) // 记录是必修课还是选修课
                {
                    $oneClassAttr = 1;
                }else 
                {
                    $oneClassAttr = 2;
                }
            }
            $oneStudentGradesInfo[$index++] = array($oneClassAttr, $oneClassCredit, $oneClassGrade); // 将每门课程得分信息存入学生分数信息数组
        }
        
        return $oneStudentGradesInfo; // 返回一个学生的所有成绩信息数组(一学期)
    }
    
    // 将该名学生一学年里两学期的所有分数累加到一个数组中存储
    function getOneStudentTwoTermGrades(&$id, &$name, $oneStudentSeq, $sheet1, $sheet2) 
    {
        $id = $sheet1->getCell("B" . $oneStudentSeq)->getValue(); // 学生的学号
        $name = $sheet1->getCell("C" . $oneStudentSeq)->getValue(); // 学生的姓名
        $oneStudentGradesInfo1 = getOneStudentGradesInfo($oneStudentSeq, $sheet1);
        $oneStudentGradesInfo2 = getOneStudentGradesInfo($oneStudentSeq, $sheet2);
        $oneStudentTwoTermGrades = array_merge($oneStudentGradesInfo1, $oneStudentGradesInfo2); // 两个成绩数组合并为1个

        return $oneStudentTwoTermGrades;
    }

    // 返回分数结果
    function getFinalResult($gradesInfo)
    {
        $compulsoryCredits = 0.0; // 必修课总学分
        $selectiveCredits = 0.0; // 选修课总学分
        
        $compulsoryGrades = 0.0; // 必修课总分数
        $selectiveGrades = 0.0; // 选修课总分数

        $finalGrades = 0.0; // 最终的分数

        for ($i = 0; $i < count($gradesInfo); $i++)
        {
            // 分别根据必修课和选修课来按学分累加必修课分数、选修课分数以及必修课学分、选修课学分
            if ($gradesInfo[$i][0] == 1)
            {
                $compulsoryCredits += $gradesInfo[$i][1];
                $compulsoryGrades += $gradesInfo[$i][2] * $gradesInfo[$i][1];
            }
            else
            {
                $selectiveCredits += $gradesInfo[$i][1];
                $selectiveGrades += $gradesInfo[$i][2] * $gradesInfo[$i][1];
            }
        } 

        // 最终成绩为必修课总分数除以必修课总学分, 选修课总分数除以选修课总学分, 分别按一定比例相加
        $finalGrades = (($compulsoryGrades / $compulsoryCredits) * 0.7 + ($selectiveGrades / $selectiveCredits) * 0.3) * 0.9;


        return $finalGrades;
    }

    // 将最终获得的成绩信息写入excel中
    function writeResultToExcel($result)
    {
        $objPHPExcel = new PHPExcel(); // 新建一个excel对象
        $objPHPExcel->setActiveSheetIndex(0); // 设置该excel对象的当前活动表格

        $column = "A"; // 列从A开始, A为列的基下标
        $row = 1; // 行下标从1开始
        // 循环遍历该二维结果数组, 将每个值写入excel中的指定的每个单元格中
        foreach ($result as $item)
        {
            $j = 0; // 列的变下标
            foreach ($item as $key => $value)
            {
                // 用ord函数返回单个字符的ascii码对应的整数值, 加1后再使用chr函数转换为单个字符
                $objPHPExcel->getActiveSheet()->setCellValue(chr(ord($column)+$j) . $row, $value);
                $j++;
            }
            $row++;
        }

        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel); // 用该excel对象来创建一个excel2007写入器
        $objWriter->save(str_replace('.php', '.xlsx', __FILE__)); // 保存该excel文件
    }

    function main($fileName1, $fileName2)
    {
        $objPHPExcel1 = openExcelFile($fileName1);
        $objPHPExcel2 = openExcelFile($fileName2);

        $sheet1 = $objPHPExcel1->getSheet(0); // 获取第一张excel中的第一个表格
        $sheet2 = $objPHPExcel2->getSheet(0); // 获取第二张excel中的第一个表格

        $allRow = $sheet1->getHighestRow(); // 获取表格的所有行数, 通常情况下, 两个表格中的行数应相等, 所以获取第一个表格中的行数即可

        $id = ""; // 学号临时变量
        $name = ""; // 姓名临时变量
        $allStudentsGradesInfo = array(); // 存放最终所有学生的分数信息数组
        $index = 0; // 存放每个学生最终分数信息的数组下标

        for ($row = 4; $row <= $allRow; $row++)
        {
            $grades = getFinalResult((getOneStudentTwoTermGrades($id, $name, $row, $sheet1, $sheet2)));
            $allStudentsGradesInfo[$index]["id"] = $id;
            $allStudentsGradesInfo[$index]["name"] = $name;
            $allStudentsGradesInfo[$index]["grades"] = $grades;
            $index++;
        }

        $sortIndicator = array_column($allStudentsGradesInfo, "grades"); // 取二维数组中的grades字段作为排序字段
        array_multisort($sortIndicator, SORT_DESC, $allStudentsGradesInfo); // 利用该排序字段对原数组按分数由高到低重新排序

        foreach ($allStudentsGradesInfo as $everyStudentGradesInfo)
        {
            echo "学号: " . $everyStudentGradesInfo["id"] . " " . "姓名: " . $everyStudentGradesInfo["name"] . " " . "分数: " . $everyStudentGradesInfo["grades"] . "<br />";
        }

        writeResultToExcel($allStudentsGradesInfo); // 将成绩信息写入excel中
    }

    echo "<br /></br />";
    
    // window下文件名要进行编码转换, Linux不需要
    main(iconv("UTF-8","GBK//IGNORE", "D:\\AppServ\\www\\excel\\" . $_FILES["uploadedFile1"]["name"]),
    iconv("UTF-8","GBK//IGNORE", "D:\\AppServ\\www\\excel\\" . $_FILES["uploadedFile2"]["name"]));

?>
