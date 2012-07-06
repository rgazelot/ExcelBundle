# ExcelBundle

Service for easy Excel export with PHPExcel Library

## Instalation

### Step 1 . ( For Symfony v. < 2.1 - deps)

Add the following lines in your `deps` file :

    [ExcelBundle\]
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

Congratulations ! You have installed ExportBundle. Read the usage documentation for learn how use the service.




