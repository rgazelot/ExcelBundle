<?php

namespace Export\ExcelBundle\Services;

use PHPExcel;
use PHPExcel_Writer_Excel5;
use PHPExcel_Worksheet_Drawing;

use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;

class Export
{
    private $workbook;
    private $cursor;
    private $currentSheet;

    public function __construct()
    {
        $this->workbook = new PHPExcel();
        $this->setDefault();
    }

    /**
     *	Set default configuration of the document.
     *  This function is public to override them from the controller.
     *  @param  array  $options  Array of options
     *  @return obj    $this
     */
    public function setDefault($options = array())
    {
        $this->currentSheet = $this->workbook->getActiveSheet();

        // Default cursor cordinates
        $this->cursor = array(
            'x' => 0,
            'y' => 1,
        );

        // Default style of the sheet
        $this->currentSheet->getDefaultStyle()->applyFromArray(array(
            'font'      => array(
                'name' => 'Arial',
                'size' => 12 ,
            ),
            'alignment' => array(
                'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                'vertical'   => \PHPExcel_Style_Alignment::VERTICAL_CENTER,
            ),
            'borders'  => array(
                'allborders' => array(
                    'style'  => \PHPExcel_Style_Border::BORDER_NONE,
                )
            )
        ));

        // Set the dimensions of de cells for the Image and Title.
        $this->currentSheet->getColumnDimensionByColumn($this->cursor['x'], $this->cursor['y'])->setWidth(5);
        $this->cursor['x'] = 1;
        $this->currentSheet->getColumnDimensionByColumn($this->cursor['x'], $this->cursor['y'])->setWidth(10);
        $this->currentSheet->getRowDimension(1)->setRowHeight(30);
        $this->currentSheet->getRowDimension(2)->setRowHeight(53);

        // Set the image logo.
        //$this->importImg('balloonlogo.png');

        return $this;
    }

    /**
     *  @param  int  $sheet
     */
    public function getSheet($sheet)
    {
        $this->currentSheet = $this->workbook->getSheet($sheet);
        return $this;
    }

    /**
     *  Create new sheet
     *  @return  obj  $this
     */
    public function createSheet()
    {
        $this->currentSheet = $this->workbook->createSheet();

        return $this;
    }

    /**
     *  Set the title of the current sheet.
     *  This function is public for first sheet. Indeed, we can pass the 
     *  translation since the controller.
     *  @param  string  $title  Title of a sheet
     *  @return obj     $this
     */
    public function setNameOfSheet($title)
    {
        $title = str_replace(\PHPExcel_Worksheet::getInvalidCharacters(), '', $title);

        if (strlen($title) >= 31) {
            $title = substr($title, 0, 28);
            $last_space = strrpos($title, " ");
            $title = substr($title, 0, $last_space) . "...";
        }

        $this->currentSheet->setTitle($title);

        return $this;
    }

    /**
     *  Set the title.
     *  The style of the title is in private function styleTitle().
     *  @param  string  $data     String of data for the title
     *  @param  array   $options  Array of options
     *  @return obj     $this
     */
    public function setTitle($data, $options = array())
    {
        // Set cursor
        $this->cursor['x'] = 4;
        $this->cursor['y'] = 2;
        if (isset($options['coordinates'])) {
            $this->cursor['x'] = $options['coordinates']['x'];
            $this->cursor['y'] = $options['coordinates']['y'];
        }

        // Options merge cells.
        if (!isset($options['merge'])) {
            $nbMerge = round(strlen($data)/6, 0);
            $this->currentSheet->mergeCellsByColumnAndRow($this->cursor['x'], $this->cursor['y'], $this->cursor['x'] + $nbMerge, $this->cursor['y']);
        }

        // Options heightRow
        if (isset($options['heightRow'])) {
            $this->currentSheet->getRowDimension($this->cursor['y'])->setRowHeight($options['heightRow']);
        }

        // Options hAlignment
        if (isset($options['hAlignment'])) {
            $this->chartCustomizeCell(array('alignment' => array('horizontal' => $options['hAlignment'])));
        }

        // Set Title.
        $this->currentSheet->setCellValueExplicitByColumnAndRow($this->cursor['x'], $this->cursor['y'], $data);
        $this->chartCustomizeCell(array(
                'font' => array(
                    'bold'  => isset($options['bold']) ? $options['bold'] : true,
                    'size'  => isset($options['size']) ? $options['size'] : 25,
                    'color' => isset($options['color']) ? $options['color'] : '000000',
                ),
            )
        );

        return $this;
    }

    /**
     *  Set the statistics and their label under the title.
     *  @param  array  $data     Labels and stats
     *  @param  array  $options  Array of options
     *  @return obj    $this
     */
    public function setInfo($data, $options = array())
    {
        // Set cursor
        $this->cursor['x'] = 1;
        $this->cursor['y'] = 4;

        if (isset($options['coordinates'])) {
            $this->cursor['x'] = $options['coordinates']['x'];
            $this->cursor['y'] = $options['coordinates']['y'];
        }

        $nextWrap = $this->cursor['x'] + 8;

        foreach ($data as $label => $val) {
            $this->currentSheet->getRowDimension($this->cursor['y'])->setRowHeight(45);
            $this->chartCustomizeCell(array(
                    'font' => array(
                        'bold'  => isset($options['bold']) ? $options['bold'] : true,
                        'size'  => isset($options['size']) ? $options['size'] : 15,
                        'color' => isset($options['color']) ? $options['color'] : '000000',
                    ),
                )
            );

            $this->currentSheet->setCellValueExplicitByColumnAndRow($this->cursor['x'], $this->cursor['y'], $val);

            $this->cursor['x']++;

            $this->chartCustomizeCell(array(
                    'font' => array(
                        'bold'  => isset($options['bold']) ? $options['bold'] : true,
                        'size'  => isset($options['size']) ? $options['size'] : 15,
                        'color' => isset($options['color']) ? $options['color'] : '000000',
                    ),
                )
            );

            // Merge cells for the label.
            $this->currentSheet->mergeCellsByColumnAndRow($this->cursor['x'], $this->cursor['y'], $this->cursor['x'] + round(strlen($label)/9, 0), $this->cursor['y']);
            $this->currentSheet->setCellValueExplicitByColumnAndRow($this->cursor['x'], $this->cursor['y'], $label);

            // Cursor for next Info labels.
            $this->cursor['x'] += 3;

            if ($this->cursor['x'] == $nextWrap) {
                $this->cursor['x'] -= 8;
                $this->cursor['y'] += 2;
            }
        }

        return $this;
    }

    /**
     *  Write the fuckin table.
     *  @param  array  $data     Data
     *  @param  array  $labels   Array of labels
     *  @param  array  $options  Array of options
     *  @return obj    $this
     */
    public function writeTable($data, $labels = array(), $options = array())
    {
        // Default coordinate of the table.
        $this->cursor['x'] = 0;
        $this->cursor['y'] = 1;
        $this->currentSheet->getColumnDimensionByColumn($this->cursor['x'], $this->cursor['y'])->setWidth(5);
        $this->cursor['x']++;
        $this->currentSheet->getRowDimension($this->cursor['y'])->setRowHeight(30);
        $this->cursor['y']++;

        if (isset($options['coordinates'])) {
            $this->cursor['x'] = $options['coordinates']['x'];
            $this->cursor['y'] = $options['coordinates']['y'];
        }

        // Set correct labels
        if (isset($options['labels'])) {

            if (empty($labels)) {
                $labels = array_keys($data[0]);
            }

            foreach ($labels as $key => $label) {

                $this->chartCustomizeCell(array(
                    'font'      => array(
                        'bold'  => isset($options['labels']['bold']) ? $options['labels']['bold'] : true,
                        'size'  => isset($options['labels']['size']) ? $options['labels']['size'] : 12,
                        'color' => isset($options['labels']['color']) ? $options['labels']['color'] : 'ffffff',
                    ),
                    'fill'      => isset($options['labels']['fill']) ? $options['labels']['fill'] : '003459',
                    'alignment' => array(
                        'wrap'       => isset($options['labels']['wrap']) ? $options['labels']['wrap'] : false,
                        'horizontal' => isset($options['labels']['horizontal']) ? $options['labels']['horizontal'] : 'hcenter',
                    ),
                ));

                // Set height of the row
                $this->currentSheet->getRowDimension($this->cursor['y'])->setRowHeight((isset($options['labels']['height'])) ? $options['labels']['height'] : 25);
                $this->writeToCell($label, isset($options['mergeCols'][$key]) ? $options['mergeCols'][$key] : null);
            }
        }

        isset($options['coordinates']) ? $this->cursor['x'] = $options['coordinates']['x'] : $this->cursor['x'] = 1;
        $this->cursor['y']++;

        // Informations under labels
        if (isset($options['infos'])) {
            foreach ($data[0] as $key => $val) {
                $this->chartCustomizeCell(array(
                    'font' => array(
                        'bold'   => isset($options['infos']['bold']) ? $options['infos']['bold'] : false,
                        'size'   => isset($options['infos']['size']) ? $options['infos']['size'] : 12,
                        'italic' => isset($options['infos']['italic']) ? $options['infos']['italic'] : true,
                        'color'  => isset($options['infos']['color']) ? $options['infos']['color'] : '000000',
                    ),
                    'fill' => isset($options['infos']['fill']) ? $options['infos']['fill'] : 'eeeeee',
                ));
                // Set height of the row
                $this->currentSheet->getRowDimension($this->cursor['y'])->setRowHeight(25);
                $this->writeToCell($val, isset($options['mergeCols'][$key]) ? $options['mergeCols'][$key] : null);
            }
            array_shift($data);
            isset($options['coordinates']) ? $this->cursor['x'] = $options['coordinates']['x'] : $this->cursor['x'] = 1;
            $this->cursor['y']++;
        }

        $c = 0;
        // Write data
        foreach ($data as $line) {

            foreach ($line as $key => $col) {
                if (!isset($options['zebra'])) {
                    if ($c % 2 == 1) {
                        $this->chartCustomizeCell(array(
                            'fill' => isset($options['zebra']['color']) ? $options['zebra']['color'] : 'a0c5e3',
                        ));
                    }
                }
                $this->currentSheet->getRowDimension($this->cursor['y'])->setRowHeight(18);
                $this->writeToCell($col, isset($options['mergeCols'][$key]) ? $options['mergeCols'][$key] : null, isset($options['hAlignment'][$key]) ? $options['hAlignment'][$key] : null, isset($options['vAlignment'][$key]) ? $options['vAlignment'][$key] : null);
            }

            isset($options['coordinates']) ? $this->cursor['x'] = $options['coordinates']['x'] : $this->cursor['x'] = 1;
            $this->cursor['y']++;
            $c++;
        }

        if (isset($options['return'])) {
            return array($this->cursor['x'], $this->cursor['y']);
        }

        return $this;
    }

    /**
     *  Write the document.
     *  @param  string  $filename  The name of export
     *  @param  string  $hash      The hash which represent the name of temp folder where the export will be save
     *  @return obj     $this
     */
    public function writeExport($filename)
    {
        $writer = new PHPExcel_Writer_Excel5($this->workbook);
        $writer->save('/tmp/' . $filename . '.xls');

        return $this;
    }

    // ============== PRIVATES ============== //

    /**
     *  @param  string  $value      Value
     *  @param  int     $mergeCols  For merge cell
     *  @param  string  $hAlignment
     *  @param  string  $vAlgnment
     */
    private function writeToCell($value, $mergeCols = null, $hAlignment = null, $vAlignment = null)
    {

        $dataType = \PHPExcel_Cell_DataType::TYPE_STRING;

        // For the dateTime format
        if (is_object($value) && get_class(new \DateTime()) === get_class($value)) {
            $value = $value->format('d/m/y h:m:s');
            $this->currentSheet->getStyleByColumnAndRow($this->cursor['x'], $this->cursor['y'])->getNumberFormat()->applyFromArray(
                array('code' => \PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY)
            );
        }

        if (null !== $hAlignment) {
            $this->chartCustomizeCell(array('alignment' => array('horizontal' => $hAlignment)));
        }

        if (null !== $vAlignment) {
            $this->chartCustomizeCell(array('alignment' => array('vertical' => $vAlignment)));
        }

        $this->currentSheet->setCellValueExplicitByColumnAndRow($this->cursor['x'], $this->cursor['y'], $value, $dataType);

        // Merge options
        if (null !== $mergeCols) {
            $this->currentSheet->mergeCellsByColumnAndRow($this->cursor['x'], $this->cursor['y'], $this->cursor['x'] + ($mergeCols - 1), $this->cursor['y']);
            $this->cursor['x'] = $this->cursor['x'] + $mergeCols;
            return;
        }
        $this->cursor['x']++;
    }

    /**
     *	@param  string  $path  path of the image
     */
    private function importImg($path)
    {
        // Merge cells for the title.
        $nbMerge = 3;
        $this->currentSheet->mergeCells('B2:D2');

        $objDrawing = new PHPExcel_Worksheet_Drawing();
        $objDrawing->setPath($path);
        $objDrawing->setCoordinates("B2");
        $objDrawing->setWorksheet($this->currentSheet);
    }

    /**
     *  Chart the active cell
     *  @param  array  $options
     */
    private function chartCustomizeCell($options = array())
    {
        if (isset($options['font'])) {
            $this->currentSheet->getStyleByColumnAndRow($this->cursor['x'], $this->cursor['y'])
            ->getFont()
            ->applyFromArray(array(
                'name'   => isset($options['font']['name']) ? $options['font']['name'] : 'Arial',
                'size'   => isset($options['font']['size']) ? $options['font']['size'] : 12,
                'bold'   => isset($options['font']['bold']) ? $options['font']['bold'] : true,
                'italic' => isset($options['font']['italic']) ? $options['font']['italic'] : false,
                'color'  => array(
                    'rgb' => isset($options['font']['color']) ? $options['font']['color'] : '000000'
                )
            ));
        }

        if (isset($options['fill'])) {
            $this->currentSheet
                ->getStyleByColumnAndRow($this->cursor['x'], $this->cursor['y'])
                ->applyFromArray(array(
                    'fill' => array(
                        'type'  => \PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('rgb' => $options['fill']),
                    )
            ));
        }

        if (isset($options['alignment'])) {

            $this->currentSheet
                ->getStyleByColumnAndRow($this->cursor['x'], $this->cursor['y'])
                ->applyFromArray(array(
                    'alignment' => array(
                        'horizontal' => isset($options['alignment']['horizontal']) ? $this->getAlignment($options['alignment']['horizontal']) : 'center',
                        'vertical'   => isset($options['alignment']['vertical']) ? $this->getAlignment($options['alignment']['vertical']) : 'center',
                        'wrap'       => isset($options['alignment']['wrap']) ? $options['alignment']['wrap'] : false,
                    ),
            ));
        }
    }

    /**
     *  $param  string  $data
     */
    private function getAlignment($data) {
        switch ($data) {
            case 'left' :
                $alignment = \PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
            break;
            case 'right':
                $alignment = \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
            break;
            case 'hcenter':
                $alignment = \PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
            break;
            case 'top':
                $alignment = \PHPExcel_Style_Alignment::VERTICAL_TOP;
            break;
            case 'bottom':
                $alignment = \PHPExcel_Style_Alignment::VERTICAL_BOTTOM;
            break;
        }

        return $alignment;
    }
}
