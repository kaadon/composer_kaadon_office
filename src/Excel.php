<?php

namespace Kaadon\Office;

use Exception;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * 导出导入Excel
 *
 * Class Excel
 *
 *
 */
class Excel
{
    /**
     * 导出Excel
     *
     * @param array $list
     * @param array $header
     * @param string $filename
     * @param string $suffix
     * @param string $path
     * @return bool
     * @throws \Exception
     */
    public static function exportData(array $list = [], array $header = [], string $filename = '', string $suffix = 'xlsx', string $path = ''): bool
    {
        if (!is_array($list) || !is_array($header)) {
            return false;
        }
        // 清除之前的错误输出
        ob_end_clean();
        ob_start();
        !$filename && $filename = time();
        // 初始化
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        // 写入头部
        $hk = 1;
        foreach ($header as  $v) {
            $sheet->setCellValue(Coordinate::stringFromColumnIndex($hk) . '1', self::formatComment($v[0]));
            $hk += 1;
        }

        // 开始写入内容
        $column = 2;
        $size = ceil(count($list) / 500);
        for ($i = 0; $i < $size; $i++) {
            $buffer = array_slice($list, $i * 500, 500);

            foreach ($buffer as $k => $row) {
                $span = 1;

                foreach ($header as $key => $value) {
                    // 解析字段
                    $realData = self::formatting($value, trim(self::formattingField($row, $value[1])), $row);
                    // 写入excel
                    // 加个"\t"制表符，解决导出大数字或银行卡等在excel中被科学计数的问题
                    $sheet->setCellValue(Coordinate::stringFromColumnIndex($span) . $column, $realData . "\t");
                    $span++;
                }

                $column++;
                unset($buffer[$k]);
            }
        }

        switch ($suffix) {
            case 'xlsx':
                $writer = new Xlsx($spreadsheet);
                $contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8;";
                $extension = 'xlsx';
                break;
            case 'xls':
                $writer = new Xls($spreadsheet);
                $contentType = "application/vnd.ms-excel;charset=utf-8;";
                $extension = 'xls';
                break;
            case 'csv':
                $writer = new Csv($spreadsheet);
                $contentType = "text/csv;charset=utf-8;";
                $extension = 'csv';
                break;
            case 'html':
                $writer = new Html($spreadsheet);
                $contentType = "text/html;charset=utf-8;";
                $extension = 'html';
                break;
            default:
                throw new Exception("Unsupported file format: $suffix");
        }

        if (!empty($path)) {
            $writer->save($path);
            return true;
        } else {
            header("Content-Type: $contentType");
            header("Content-Disposition: attachment; filename=\"{$filename}.{$extension}\"");
            header("Content-extension: $extension");
            header("Content-filename: $filename");
            header('Cache-Control: max-age=0');
            $writer->save('php://output');
            exit();
        }
    }

    /**
     * 导出的另外一种形式(不建议使用)
     *
     * @param array $list
     * @param array $header
     * @param string $filename
     * @return bool
     */
    public static function exportCsvData(array $list = [], array $header = [], string $filename = ''): bool
    {
        if (!is_array($list) || !is_array($header)) {
            return false;
        }

        // 清除之前的错误输出
        ob_end_clean();
        ob_start();

        !$filename && $filename = time();

        $html = "\xEF\xBB\xBF";
        foreach ($header as $k => $v) {
            $html .= $v[0] . "\t ,";
        }

        $html .= "\n";

        if (!empty($list)) {
            $info = [];
            $size = ceil(count($list) / 500);

            for ($i = 0; $i < $size; $i++) {
                $buffer = array_slice($list, $i * 500, 500);

                foreach ($buffer as $k => $row) {
                    $data = [];

                    foreach ($header as $key => $value) {
                        // 解析字段
                        $realData = self::formatting($value, trim(self::formattingField($row, $value[1])), $row);
                        $data[] = str_replace(PHP_EOL, '', $realData);
                    }

                    $info[] = implode("\t ,", $data) . "\t ,";
                    unset($data, $buffer[$k]);
                }
            }

            $html .= implode("\n", $info);
        }

        header("Content-type:text/csv");
        header("Content-Disposition:attachment; filename={$filename}.csv");
        echo $html;
        exit();
    }

    /**
     * 导入
     *
     * @param $filePath
     * @param int $startRow
     * @return array|mixed
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function import($filePath, int $startRow = 1): mixed
    {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);
        if (!$reader->canRead($filePath)) {
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
            // setReadDataOnly Set read data only 只读单元格的数据，不格式化 e.g. 读时间会变成一个数据等
            $reader->setReadDataOnly(true);

            if (!$reader->canRead($filePath)) {
                throw new Exception('不能读取Excel');
            }
        }

        $spreadsheet = $reader->load($filePath);
        $sheetCount = $spreadsheet->getSheetCount(); // 获取sheet的数量

        // 获取所有的sheet表格数据
        $excleDatas = [];
        $emptyRowNum = 0;
        for ($i = 0; $i < $sheetCount; $i++) {
            $currentSheet = $spreadsheet->getSheet($i); // 读取excel文件中的第一个工作表
            $allColumn = $currentSheet->getHighestColumn(); // 取得最大的列号
            $allColumn = Coordinate::columnIndexFromString($allColumn); // 由列名转为列数('AB'->28)
            $allRow = $currentSheet->getHighestRow(); // 取得一共有多少行

            $arr = [];
            for ($currentRow = $startRow; $currentRow <= $allRow; $currentRow++) {
                // 从第1列开始输出
                for ($currentColumn = 1; $currentColumn <= $allColumn; $currentColumn++) {
                    $val = $currentSheet->getCellByColumnAndRow($currentColumn, $currentRow)->getValue();
                    $arr[$currentRow][] = trim($val);
                }

                // $arr[$currentRow] = array_filter($arr[$currentRow]);
                // 统计连续空行
                if (empty($arr[$currentRow]) && $emptyRowNum <= 50) {
                    $emptyRowNum++;
                } else {
                    $emptyRowNum = 0;
                }
                // 防止坑队友的同事在excel里面弄出很多的空行，陷入很漫长的循环中，设置如果连续超过50个空行就退出循环，返回结果
                // 连续50行数据为空，不再读取后面行的数据，防止读满内存
                if ($emptyRowNum > 50) {
                    break;
                }
            }

            $excleDatas[$i] = $arr; // 多个sheet的数组的集合
        }

        // 这里我只需要用到第一个sheet的数据，所以只返回了第一个sheet的数据
        $returnData = $excleDatas ? array_shift($excleDatas) : [];

        // 第一行数据就是空的，为了保留其原始数据，第一行数据就不做array_fiter操作；
        $returnData = $returnData && isset($returnData[$startRow]) && !empty($returnData[$startRow]) ? array_filter($returnData) : $returnData;
        return $returnData;
    }

    /**
     * 格式化内容
     *
     * @param array $array 头部规则
     * @return false|mixed|null|string 内容值
     */
    protected static function formatting(array $array, $value, $row): mixed
    {
        $formatDefine = [];
        preg_match('/\([\s\S]*?\)/i', $array[0], $formTypeMatch);
        if (!empty($formTypeMatch) && isset($formTypeMatch[0])) {
            $formType = trim(str_replace(')', '', str_replace('(', '', $formTypeMatch[0])));
            if (isset($formType) && in_array($formType, ['select', 'selects', 'radio', 'checkbox', 'switch'])) {
                $arr = explode(":", $array[0]);
                if (isset($arr[1])) {
                    $optionArr = explode(",", $arr[1]);
                    if ($optionArr) {
                        foreach ($optionArr as $k => $v) {
                            $dataOption = explode("=", $v);
                            if (isset($dataOption[0]) && isset($dataOption[1])) {
                                $formatDefine[trim($dataOption[0])] = trim($dataOption[1]);
                            }
                        }
                    }

                }
            }
        }
        $formatData = "";
        if (!empty($formatDefine)) {
            $valueArr = explode(',', $value);
            $i = 0;
            foreach ($valueArr as $v) {
                $formatData = $i == 0 ? $formatDefine[$v] : $formatData . "," . $formatDefine[$v];
                $i++;
            }
        }
        return $formatData ?: $value;
    }

    /**
     * 解析字段
     *
     * @param $row
     * @param $field
     * @return mixed
     */
    protected static function formattingField($row, $field): mixed
    {
        $newField = explode('.', $field);
        if (count($newField) == 1) {
            return $row[$field] ?? '-';
        }
        return $row[$newField[0]][$newField[1]] ?? '-';
    }

    /**
     * 格式化备注
     */
    protected static function formatComment($value)
    {
        $comment = $value;
        $arr = explode(":", $value);
        if (isset($arr[0])) {
            $comment = $arr[0];
            $comment = str_replace("(selects)", '', $comment);
            $comment = str_replace("(select)", '', $comment);
            $comment = str_replace("(radio)", '', $comment);
            $comment = str_replace("(checkbox)", '', $comment);
        }
        return $comment;
    }
}
