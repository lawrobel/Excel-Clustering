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
Private Function discrete_random_variable(ByRef x_values As Variant, ByRef probabilities As Variant) As Integer
    '#This function simulates a discrete random variable from an array of values and an array of probabilities
    '#The probability of the ith entry of x_values is the ith entry of probabilities
    '#Returns a random value of x_values according to the probability distribition given by probabilities
    '#Assumes n > 2
    Dim n As Double
    Dim rolling_sum As Double
    Dim rand_zero_one As Double 'random number between zero and one
    n = UBound(x_values)
    ReDim cumulative_probabilities(1 To n) As Variant 'divides up the interval 0 to 1 into bins
    
    rolling_sum = 0
    For i = 1 To n
        rolling_sum = rolling_sum + probabilities(i)
        cumulative_probabilities(i) = rolling_sum 'cumulative probabilities is sum of all previous probabilities
    Next i
    rand_zero_one = Rnd()
    
    ' the code below chooses the x value based on what bin the probability of the x value falls into
    If rand_zero_one < cumulative_probabilities(1) Then
        discrete_random_variable = x_values(1)
    End If
    For i = 2 To (n - 1)
        If rand_zero_one > cumulative_probabilities(i) And rand_zero_one < cumulative_probabilities(i + 1) Then
            discrete_random_variable = x_values(i)
        End If
    Next i
    If rand_zero_one > cumulative_probabilities(n - 1) Then
        discrete_random_variable = x_values(n)
    End If
    
End Function
Private Sub k_means_pp(ByRef x As Variant, ByRef centroids As Variant, k As Variant, xdim As Variant, xnumpoints As Variant)
    '# The sub initializes cluster centers by using the k-means++ algorithm
    '# The sub takes in x array (input range as an array), centroids array, k parameter, xdim and xnumpoints which are all integers
    '# but are these last three are passed in as variant types to avoid errors
    
    Dim random_int As Integer 'random index to determine point in input range to select
    Dim sum_of_squared_distances As Double
    
    ReDim random_point(1 To xdim) As Variant 'actual point from input range determined by random_int
    ReDim data_point(1 To xdim) As Variant
    ReDim centroid(1 To xdim) As Variant
    ReDim min_squared_distances(1 To xnumpoints) As Variant
    ReDim probability_array(1 To xnumpoints) As Variant
    ReDim index_array(1 To xnumpoints) As Variant
    
    random_int = CInt(WorksheetFunction.Ceiling_Math(Rnd() * xnumpoints)) + 1 'rand between 1 and xnumpoints
    
    'the below loop does the same thing as Application.Index(x, random_int, 0) but there were issues with that here so this is a equivalent way
    For m = 1 To xnumpoints
        index_array(m) = m 'for use later in discrete random variable function
        If m = random_int Then
            For l = 1 To xdim
                random_point(l) = x(random_int, l)
            Next l
        End If
    Next m
    
    For p = 1 To xdim
        centroids(1, p) = random_point(p) 'first choose random center from among the datapoints
    Next p
        
    For i = 2 To k
        sum_of_squared_distances = 0
        ReDim distances(1 To i - 1) As Variant 'only consider distances from points to already chosen centroids, this is why distances gets redefined in the loop
        For j = 1 To xnumpoints
            data_point = Application.Index(x, j, 0)
            For p = 1 To (i - 1)
                centroid = Application.Index(centroids, p, 0)
                distances(p) = distance(data_point, centroid)
            Next p
            min_squared_distances(j) = Application.Min(distances) ^ 2
            sum_of_squared_distances = sum_of_squared_distances + min_squared_distances(j)
        Next j
            
        For a = 1 To xnumpoints
            probability_array(a) = min_squared_distances(a) / sum_of_squared_distances 'define probability distribution
        Next a
        random_int = discrete_random_variable(index_array, probability_array)
        For m = 1 To xnumpoints 'since there was issues with using application.index
            If m = random_int Then
                For l = 1 To xdim
                    random_point(l) = x(random_int, l)
                Next l
            End If
        Next m
        For p = 1 To xdim
            centroids(i, p) = random_point(p) 'first choose random center from among the datapoints
        Next p
    Next i
End Sub

Public Sub k_means()
    '# This macro runs the kmeans clustering algorithm on a two-dimensional range of data.
    '# The value of k is determined by the value of cell M2 and the maximum number of iterations
    '# is determined by cell M3. The centers are either chosen by kmeans++ or randomly selected to start.
    '# The center initialization method is specified by the user in cell O3.
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
    Dim xdim As Variant
    Dim xnumpoints As Variant
    
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
    
    
    If Range("O3").Value = "random" Then
        For i = 1 To k
            For j = 1 To xdim
                columns_array = Application.Index(x, 0, j) 'jth column of two-dim array x
                max_value = Application.Max(columns_array)
                min_value = Application.Min(columns_array)
                centroids(i, j) = min_value + Rnd() * (max_value - min_value)
            Next j
        Next i
    
    ElseIf Range("O3").Value = "k-means++" Then
        Call k_means_pp(x, centroids, k, xdim, xnumpoints) 'runs the k_means_pp private subroutine
    End If
    
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
        DoEvents
    Loop
    Range("M17").Value = count_it
    
    For i = 2 To (xnumpoints + 1)
         Range("D" & i).Value = clusters(i - 1) ' put final cluster labels next to each point in input range
    Next i
    
End Sub
