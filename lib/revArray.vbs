option explicit

'test_revArray

sub revArray(a,k,n)
	dim i, t(10000)
	for i=1 to n
		t(i)=a(k,i)
	next
	for i=1 to n
		a(k,i)=t(n-i+1)
	next
end sub

sub test_revArray
	const n=5, k=3
	dim x(20,100),i,s
	for i=1 to n
		x(k,i)=i
	next
	revArray x,k,n
	s=""
	for i=1 to n
		s = s & " " & x(k,i)
	next
	msgbox s
end sub
