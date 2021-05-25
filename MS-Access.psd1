@{
   RootModule        = 'MS-Access.psm1'
   ModuleVersion     = '0.1'
   RequiredModules   = @(
      'COM',
      'vb'
   )
   FunctionsToExport = @(
     'get-msAccess'                ,
     'show-msAccess'               ,
     'invoke-msAccessQuery'        ,
     'invoke-msAccessQueryFromFile'
   )
   AliasesToExport   = @(
     'acc-query'                   ,
     'acc-queryFromFile'
   )
}
