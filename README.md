# CubeRetriever

This project provides an easy API for connecting to a MS Analysis Services OLAP cube / SSAS cube and executing MDX queries.

If you are using Microsoft SQL Server Management Studio to run MDX queries and want to further manipulate the data returned with Pandas or Matplotlib. Or if you wish to store the data directly into an Excel file. This package can help you.

Tested on python==3.5.6

## 1. Install Microsoft R Client

Microsoft R Client ships with olapR package that is required for connecting to cube.

    https://docs.microsoft.com/en-us/machine-learning-server/r-client/install-on-windows

## 2. Install packages required

If conda is used as envirement management tool, make sure to disable the installation of dependencies($ conda install xxx --no-deps). Otherwise, conda would install its own version of R runtime.

    $ pip install -r requirements.txt
    
## 3. Install openxlsx Package in R

Optional. Only if you want to directly store the data returned from cube to excel files.

    >>> import rpy2.robjects.packages as r
    >>> utils = r.importr("utils")
    >>> package_name = "openxlsx"
    >>> utils.install_packages(package_name)

## 4. Change the cube connection configuration in config.ini file

Change the default data_source and initial_catalog according to your own cube setup.

    data_source=cubes.abc.com
    provider=MSOLAP
    initial_catalog=ABC Group

## 5. Examples

Execute MDX queries and get the return data in the format of Pandas Dataframe.

    >>> import cube_retriever as cube
    >>> retriever = cube.CubeRetriever() 
    >>> retriever.conn() # Establish connection to cube
    >>> mdx_qry_1 = """
            SELECT  
            { [Measures].[Sales Amount],   
                [Measures].[Tax Amount] } ON COLUMNS,  
            { [Date].[Fiscal].[Fiscal Year].&[2002],   
                [Date].[Fiscal].[Fiscal Year].&[2003] } ON ROWS  
            FROM [Adventure Works]  
            WHERE ( [Sales Territory].[Southwest] )  
        """
    >>> mdx_qry_2 = """
            SELECT
            { [Measures].[Unit Sales], [Measures].[Store Sales] } ON COLUMNS,
            { [Time].[1997], [Time].[1998] } ON ROWS
            FROM Sales
            WHERE ( [Store].[USA].[CA] )
        """
    >>> df_1 = retriever.mdx_query(mdx_qry_1)
    >>> df_2 = retriever.mdx_query(mdx_qry_2)

Execute one line of MDX query and store the data into an excel file.

    >>> import cube_retriever as cube
    >>> mdx_qry = """
            SELECT  
            { [Measures].[Sales Amount],   
                [Measures].[Tax Amount] } ON COLUMNS,  
            { [Date].[Fiscal].[Fiscal Year].&[2002],   
                [Date].[Fiscal].[Fiscal Year].&[2003] } ON ROWS  
            FROM [Adventure Works]  
            WHERE ( [Sales Territory].[Southwest] )  
        """
    >>> output_file = "abc.xlsx"
    >>> retriever = cube.CubeRetriever()
    >>> retriever.mdx_query_to_excel(mdx_qry, output_file)

