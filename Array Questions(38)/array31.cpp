#include <bits/stdc++.h>
using namespace std;
int brute_largest_zero_sum_subarray(vector<int>& a){
    int n=a.size(), best=0;
    for(int i=0;i<n;i++){ int s=0; for(int j=i;j<n;j++){ s+=a[j]; if(s==0) best=max(best, j-i+1); } }
    return best;
}
int optimal_largest_zero_sum_subarray(vector<int>& a){
    unordered_map<int,int> first;
    int s=0, best=0;
    first[0]=-1;
    for(int i=0;i<a.size();++i){ s+=a[i]; if(first.find(s)==first.end()) first[s]=i; else best=max(best, i-first[s]); }
    return best;
}
int main(){return 0;}