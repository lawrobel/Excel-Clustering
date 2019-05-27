Attribute VB_Name = "Module1"
Private Function distance(ByRef point1 As Variant, ByRef point2 As Variant) As Double
    '# This function computes the euclidean distance between the two given points (arrays)
    
    Dim sum As Double
    sum = 0
    For i = LBound(point1) To UBound(point1) '# point 1 and point 2 are assumed to have same size
        sum = sum + (point1(i) - point2(i)) ^ 2
    Next i
    distance = Sqr(sum)
    
End Function

Public Sub k_means()
    '# This macro runs the kmeans clustering algorithm on a two-dimensional range of data.
    '# The value of k is determined by the value of cell M2 and the maximum number of iterations
    '# is determined by cell M3. The centers are randomly selected to start.
    '# The total error between the centroids and the previous centroids is placed in row M16 and
    '# updates every iteration. The number of points in each cluster as well as the cluster centers
    '# are placed on the sheet as well in the range F1:J5; these are both also updated at each iteration.
    
    Dim k As Integer
    Dim max_iter As Integer
    Dim num_rows As Integer 'number of rows in input range
    Dim max_value As Double 'max value of specific variable
    Dim min_value As Double
    Dim x As Variant 'to contain the input range as an array
    Dim total_error As Double
    Dim cluster As Integer 'for labels for each point
    Dim error_tol As Double 'user specified
    Dim count_it As Integer 'to count the number of iterations of the while loop
    
    k = Range("M2").Value 'user specified on sheet, keep default values in cells like this
    num_rows = Range("M18").Value
    max_iter = Range("M3").Value 'user specified on sheet
    
    Set input_range = Range("B2:C" & num_rows) 'data to be clustered
    xnumpoints = input_range.Rows.Count
    xdim = input_range.Columns.Count

    x = input_range.Value 'turn range into array
    
    ReDim centroids(1 To k, 1 To xdim) As Variant
    ReDim previous_centroids(1 To k, 1 To xdim) As Variant
    ReDim columns_array(1 To xnumpoints) As Variant
    
    For i = 1 To k
        For j = 1 To xdim
            columns_array = Application.Index(x, 0, j) 'jth column of two-dim array x
            max_value = Application.Max(columns_array)
            min_value = Application.Min(columns_array)
            centroids(i, j) = min_value + Rnd() * (max_value - min_value)
        Next j
    Next i
    
    For i = 2 To (k + 1) 'needs to start at 2 because of the header string
        Range("I" & i).Value = centroids(i - 1, 1)
        Range("J" & i).Value = centroids(i - 1, 1)
    Next i
    
    ReDim clusters(1 To xnumpoints) As Variant
    ReDim errorarray(1 To k) As Variant
    
    total_error = 0
    For i = 1 To k
        ReDim arr1(1 To k) As Variant
        ReDim arr2(1 To k) As Variant
        arr1 = Application.Index(centroids, i, 0)
        arr2 = Application.Index(previous_centroids, i, 0)
        errorarray(i) = distance(arr1, arr2)
        total_error = total_error + errorarray(i)
    Next i
    
    Range("M16").Value = total_error 'see initial error
    previous_centroids = centroids
    
    ReDim distances(1 To k) As Variant
    ReDim temparr3(1 To k) As Variant
    ReDim size_of_clusters(1 To k) As Variant 'for the number of points in each cluster, needed for average value
    
    error_tol = Range("O2").Value
    
    Do While total_error > error_tol And count_it < max_iter 'can change error tolarance later
        
        ReDim size_of_clusters(1 To k) As Variant  'set back to zero
        
        For i = 1 To xnumpoints
            ReDim distances(1 To k) As Variant 'set back to zero
            For j = 1 To k
                temparr1 = Application.Index(x, i, 0)
                temparr2 = Application.Index(centroids, j, 0)
                distances(j) = distance(temparr1, temparr2)
            Next j
            
            cluster = Application.Match(Application.Min(distances), distances, 0) 'argmin
            clusters(i) = cluster 'assign point to clostest cluster
            
            For p = 1 To k 'keep track of size of clusters for averaging later
                If cluster = p Then
                    size_of_clusters(p) = size_of_clusters(p) + 1
                End If
            Next p
        Next i

        For i = 2 To (k + 1)
            Range("G" & i).Value = size_of_clusters(i - 1) 'see cluster sizes as algorithm runs
        Next i
        
        ReDim cluster_points(1 To xnumpoints, 1 To xdim) As Variant
        ReDim sums(1 To k, 1 To xdim) As Variant 'need to sum each cluster and each dimension of cluster
        ReDim temparr4(1 To k) As Variant
        
        For i = 1 To k
            For j = 1 To xnumpoints
                If clusters(j) = i Then ' add to sum when point j is in cluster j
                    temparr4 = Application.Index(x, j, 0)
                    For n = 1 To xdim
                        cluster_points(j, n) = temparr4(n)
                        sums(i, n) = sums(i, n) + cluster_points(j, n)
                    Next n
                End If
            Next j
            
            For m = 1 To xdim
                If size_of_clusters(i) = 0 Then 'avoids issues with dividing by zero
                    centroids(i, m) = 0
                Else:
                    centroids(i, m) = sums(i, m) / size_of_clusters(i) 'average value of cluster i in dimension m
                End If
            Next m
        Next i
        
        For i = 2 To (k + 1)
            Range("I" & i).Value = centroids(i - 1, 1) 'see the centroid coordinates as the algorithm runs
            Range("J" & i).Value = centroids(i - 1, 2)
        Next i

        total_error = 0 'set error back to zero since it was set before
        For i = 1 To k
            temparr1 = Application.Index(centroids, i, 0)
            temparr2 = Application.Index(previous_centroids, i, 0)
            errorarray(i) = distance(temparr1, temparr2)
            total_error = total_error + errorarray(i)
        Next i
        previous_centroids = centroids
        Range("M16").Value = total_error 'see how the error changes as the algorithm runs

        count_it = count_it + 1
    Loop
    
    For i = 2 To (xnumpoints + 1)
         Range("D" & i).Value = clusters(i - 1) ' put final cluster labels next to each point in input range
    Next i
    
End Sub
