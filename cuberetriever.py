"""
Cube Retriever

Package for executing mdx queries and extracting data from cube.
"""
from configparser import ConfigParser
import os
from rpy2.robjects.packages import importr
from rpy2.robjects import pandas2ri
import rpy2.robjects as robjects


class CubeRetriever(object):
    """
    Retrieve data from cube.
    """
    def __init__(self, ini_file=os.path.join(os.path.abspath(os.path.dirname(__file__)), 'config.ini')):
        """
        Initialization, connection settings
        """
        config = ConfigParser()
        config.read(ini_file)
        self.data_source = config.get('connection', 'data_source')
        self.provider = config.get('connection', 'provider')
        self.initial_catalog = config.get('connection', 'initial_catalog')
        self.connstr = "Data Source=" + self.data_source + "; Provider=" + self.provider + "; initial catalog=" + self.initial_catalog
        self.olapr = None
        self.olapCnn = None

    def conn(self):
        """
        Establish connection to cube
        """
        self.olapr = importr('olapR')
        self.olapCnn = self.olapr.OlapConnection(self.connstr)

    def mdx_query(self, mdx_query):
        """Make a mdx query in cube, returns a Pandas DataFrame

        Might result in error if there are characters in rare encoding.
        When there is an error, try method mdxQuery_to_excel()

        Args:
        - mdx_query: String of mdx query
        Return:
        - Pandas DataFrame of results
        """
        res = self.olapr.execute2D(self.olapCnn, mdx_query)
        res_df = pandas2ri.ri2py_dataframe(res)
        return res_df

    def mdx_query_to_excel(self, mdx_query, filename, append="FALSE"):
        """Make a mdx query and write the resulting dataframe into an excel file

        Args:
            mdx_query, string of MDX qurey language
            filename, string of filename(excel) to be stored
        """
        # escape double quote
        mdx_query = mdx_query.replace('"', '\\\"')

        robjects.r('library(olapR)')

        # xlsx is Java-based, 'out of memory' error occurs when DF returned is too large
        # robjects.r('library(xlsx)')

        # Using openxlsx instead, which is CPP-based
        robjects.r('library(openxlsx)')

        # Connect and query
        robjects.r('cnnstr <- "Data Source=' + self.data_source + '; Provider=' + self.provider + '; initial catalog=' + self.initial_catalog + '"')
        robjects.r('olapCnn <- OlapConnection(cnnstr)')
        robjects.r('query_str="' + mdx_query + '"')
        robjects.r('result <- execute2D(olapCnn, query_str)')

        # Write into excel file
        robjects.r('filename <- "' + filename + '"')
        robjects.r('write.xlsx(result, filename, sheetName = "Sheet1", col.names = TRUE, row.names = TRUE, append = ' + append + ')')

    def explorer(self, *argv):
        """
        Print out all members of the specified dimension and hierarchy
        """
        self.olapr.explore(self.olapCnn, *argv)
