# Usage Export

## Presentation

This service is responsible for processing the excel export. It has many functions which purpose is to create a sheet, a title, informations, tables, etc... This documentation explain any functions and stuff good to know of the service.

The goal of the service is you never need to use the functions of PHPExcel library.

This service is accessible anywhere in your application. Just use the `get()` function in your Controller to access it, like this : `$this->get('export.excel')`.

### Table fo content

1. Create a sheet
2. Set name of a sheet
3. Set title
4. Write a table
5. Import image
6. Get sheet
7. Write the export
8. Foreword
9. Other important stuff

## 1. Create a sheet

You can create a lot of sheet with the function `createSheet()`. This function set automatically the current sheet for the new sheet recently created.

#### Important :

When the service is call with `get('export.excel')`, the `setDefault()` function create the first sheet. So you do not have create.

## 2. Set name of a sheet

The name of a sheet is displayed on the tab in Excel. You can set it with the function `setNameOfSheet()`.

#### Parameters :

	1. $title : string

## 3. Set title

The `setTitle()` function create a little table in one cell and with title font by default. By default, you have the following font : `25 bold #000000` (where 25 is the size). But you can change anything with the options in parameters.

#### Parameters :

	1. $data : string // The title
	2. $options : array // The options

	$options = array(
		'bold' => true/false,
		'size' => int,
		'color' => hexa,
		'merge' => true/false,
		'coordinates' => array(
			'x' => int,
			'y' => int,
		),
		'heightRow' => int,
		'hAlignment' => string : left/right/center,
	)

#### Options details :

* `merge` : if you set merge at true, the number of merge depend of the length of the title string.
* `coordinates` : you can define the abscissa and ordinate where the title will be write.
* `heightRow` : the height of the row.
* `hAlignment` : the horizontal text style in the cell.

## 4. Write a table

The function `writeTable` is probably the most important function in the service. It can help you to drawing a lot of table with its multiple options.

#### Parameters :

	1. $data : array // All the data preformatted
	2. $labels : array
	3. $options : array

	$data = array(
		array(
			'user' => Thomas,
			'id'   => 23,
			'age'  => 19,
			'team' => France,
		),
		array(
			'user' => Jack,
			'id'   => 34,
			'age'  => 22,
			'team' => France,
		),
		.
		.
		.
	)

Each arrays represent a row.

	$labels = array(
		'user' => Users,
		'id'   => ID,
		'age'  => Age,
		'team' => Team,
	)

Each indexes of the array `$label` are the same in the `$data`. They represent the name of the column. You can also use the translator service for give the good translation for each labels.

	$options = array(
		'coordinates' => array(
			'x' => int,
			'y' => int,
		),
		'labels' => array(
			'bold'       => true/false,
			'size'       => int,
			'color'      => hexa,
			'fill'       => hexa,
			'height'     => int,
			'wrap'       => true/false,
			'horizontal' => string : left/right/center,
		),
		'mergeCols' => array(
			'user' => int,
			'id'   => int,
			'age'  => int,
			.
			.
		),
		'infos' => array(
			'bold'   => true/false,
			'size'   => int,
			'italic' => true/false,
			'color'  => hexa,
			'fill'   => hexa,
		),
		'hAlignment' => array(
			'user' => string : left/right/center,
			'id'   => string : left/right/center,
			'age'  => string : left/right/center,
			.
			.
			.
		),
		'vAlignment' => array(
			'user' => string : top/bottom,
			'id'   => string : top/bottom,
			'age'  => string : top/bottom,
			.
			.
			.
		),
		'zebra' => array(
			'color' => hexa,
		),
		'return' => true/false,
	)

#### Options details :

* All options are optional but beware options have a default comportment, see the following lines.
* `zebra` : by default zebra is active with a default color. You can change it with the option `color`. If you want disable the zebra, you must set the option at `'zebra' => false`.
* `infos` : by default, informations are disable. If you want active them and use the defaults styles, set `'infos' => true`.
* `labels` : by default, the labels are disable. If you want active them and use the defaults styles, set `'labels' => true`.
* `return` : by default, return is disable. If you activate it, the `writeTable()` function will return the last coordinates where it stopped. It's very useful if you want write lot of table in the same sheet.
* `mergeCols` : by default, the merge is one cell.
* `coordinates` : the default coordinates are x = 0 and y = 1.

## 5. Import Image

The goal of this function is import image where you want in the sheet. this function is excute with `importImg()`.

#### Parameters :

	1. $path : sting
	2. $options : array

	$options = array(
		'coordinates'    => array(
			'x' => int,
			'y' => int,
		),
		'merge'          => int,
		'imgCoordinates' => string,
		'heightRow'      => int,
	)

#### Options details :

require :
* `imgCoodinates` : example 'A3', 'B4' or 'H5'.

optional :

* `coordinates` : the default coordinates are x = 0 and y = 1.
* `merge` : by default, the merge is one cell.

## 6. Get sheet

Maybe you want access to a specific sheet for set a summary, or write another table. You can make it with the function `getSheet()`. Just pass the number of the sheet what you want and the function sets the current sheet to this sheet.

#### Parameters :

	1. $sheet : int

## 7. Write the export

The export must be write and launch with the function `writeExport()`. This function use `PHPExcel_Writer_Excel5` for create a file in the tmp. The extension of the export will be .xls.

#### Parameters :

	1. $filename : string

## 8. How the service stylize a cell ?

For stylize a cell, you need to use different functions of PHPExcel library like `getFont()`, `getStyle()` or `getAlignment()`. All of those functions are embedded in a single function named `chartCustomizeCell()`. It is a private function of the service and you can access it simply with `$this->chartCustomizeCell()` evrywhere in the service. You must stylize all cells with this function.

#### Parameters :

The function take an array for options. Here you have all the differents styles possible with the function :

	1.  $options = array(
			'font' => array(
				'name'   => string,
				'size'   => int,
				'bold'   => true/false,
				'italic' => true/false,
				'color'  => array(
					'rgb' => hexa,
				),
			),
			'fill' => array(
				'color' => array(
					'rgb' => hexa,
				),
			),
			'alignment' => array(
				'horizontal' => string : left/right/center,
				'vertical'   => string : left/right,
				'wrap'       => true/false,
			),
		)

## 9. Other important stuff

#### Chainability

ALl the function of the service can be chain. You can create a query like this :

	$this->get('export.excel')
		 ->export
         ->createSheet()
         ->setNameOfSheet($this->translator->trans('export_quotes_sheet'))
         ->setTitle($this->translator->trans('export_quotes_sheet'), array(
            'coordinates' => array('x' => 1, 'y' => 2),
            'size'        => 14,
            'merge'       => false,
            'hAlignment'  => 'left',
            'heightRow'   => 53,
         ))
         ->writeTable($arrayQuotes, $arrayLabels, array(
            'mergeCols'   => array(
                'users'     => 2,
                'quotes'    => 7,
                'popupSet'  => 2,
                'popupTime' => 2,
                'sessions'  => $mergeSessionCell,
            ),
            'hAlignment'  => array(
                'quotes' => 'left',
                'users'  => 'left',
            ),
            'coordinates' => array('x' => 1, 'y' => 5),
            'labels'      => true,
            'infos'       => true,
         ));

#### Developpers

This service was wrote by Rémy Gazelot : Github : [rgazelot](https://github.com/rgazelot) / Twitter : [@remygazelot](https://twitter.com/#!/remygazelot)


















