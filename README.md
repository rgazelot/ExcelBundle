# Deprecated bundle

This library is deprecated. I suggest you to use this new library write to easily build table and handle PHPExcel : https://github.com/Wisembly/ExcelAnt

# ExcelBundle

Service for easy Excel export with PHPExcel Library

## Installation

### Step 1 . ( For Symfony v. < 2.1 - deps)

Add the following lines in your `deps` file :

    [ExcelBundle]
        git=http://github.com/rgazelot/ExcelBundle.git
        target=/bundles/Export/ExcelBundle

Now download the bundle by running the command :

    ./bin/vendors update

Symfony will install your bundle in `vendors/bundles/Exports`

Declare in the `autoload.php` :

    # app/autoload.php

    $loader->registerNamespaces(array(
        ...
        'Export' => __DIR__.'/../vendor/bundles',
    ));

### Step 1 . ( For Symfony v. >= 2.1 - with composer)

coming soon ...

### Step 2 . Enable your bundle

Enable the bundle in the kernel :

    # app/AppKernel.php

    $bundles = array(
        ...
        new Export\ExcelBundle\ExportExcelBundle(),
    );

### Step 3 . Register PHPExcel with Prefixes

Because the PHPExcel library not use namespaces, create prefixes in `autoload.php`.

    # app/autoload.php

    $loader->registerPrefixes(array(
        ...
        'PHPExcel' => __DIR__.'/../vendor/bundles/Export/ExcelBundle/Library/phpExcel/Classes',
    ));

### Step 4 . Configure your bundle

In your configuration file :

    # app/config/config.yml

    export_excel: ~

### Finish !

Congratulations ! You have installed ExportBundle. Read the [usage documentation](https://github.com/rgazelot/ExcelBundle/blob/master/USAGE.md) for learn how to use the service.




