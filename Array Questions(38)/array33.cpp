#include <bits/stdc++.h>
using namespace std;
void brute_merge_two_sorted_inplace(vector<int>& a, int m, vector<int>& b, int n){
    vector<int> c;
    for(int i=0;i<m;++i) c.push_back(a[i]);
    for(int j=0;j<n;++j) c.push_back(b[j]);
    sort(c.begin(), c.end());
    a=c;
}
void optimal_merge_two_sorted_inplace(vector<int>& a, int m, vector<int>& b, int n){
    int i=m-1, j=n-1, k=m+n-1;
    while(j>=0){
        if(i>=0 && a[i]>b[j]) a[k--]=a[i--];
        else a[k--]=b[j--];
    }
}
int main(){return 0;}