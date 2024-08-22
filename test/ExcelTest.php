<?php

namespace Kaadon\Test;

use Kaadon\Office\Excel;
use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase
{
    public function testExportData_withValidDataAndXlsxSuffix_shouldReturnTrue(): void
    {
        $data = json_decode(file_get_contents(__DIR__ . '/data.json'), true);
        list("list" => $list, "header" => $header, "filename" =>$filename) = $data;
        $suffix = 'xlsx';
        $path = __DIR__  . $filename . '.' . $suffix;
        $result = Excel::exportData($list, $header, $filename, $suffix, $path);
        $this->assertTrue($result);
    }
}