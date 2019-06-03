# Excel-Clustering
Clustering Algorithms Implemented in Microsoft Excel

## What is this Workbook and What is its Use?

<p>This Excel workbook implements the k-means clustering algorithm and the k-means++ algorithm on a range of data 
containing two columns and any number of rows. VBA was used to implement the algorithms as well as write macros 
for simulating data, plotting, coloring and formatting. This workbook provides a simple way to cluster small
datasets and visualize the results within the spreadsheet itself.</p>

## Structure and Content of the Workbook

<p>Within the 2-D Clustering sheet, the user can first simulate clustered data using a button or they can copy and
paste their own data into the two data columns. The user can then setup the worksheet to the parameters they want to use such
as the k parameter and centroid initialization method (random or k-means++). Then they can press a button to run the k-means algorithm. After the algorithm finishes, the centroids are plotted and the cluster ID is given to each data point in the column next to the data.</p>

<p>Other algorithm results are given in the sheet such as how much data points there are per cluster and where the centroids are located. The clusters can also be colored using the 'Color Clusters' button. This allows the user to visualize which data points belong to the same cluster.</p>
