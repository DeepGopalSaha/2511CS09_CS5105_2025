#include <bits/stdc++.h>
using namespace std;
int brute_longest_subarray_sum_k(vector<int>& a,int k){
    int n=a.size(), best=0;
    for(int i=0;i<n;++i){ int s=0; for(int j=i;j<n;++j){ s+=a[j]; if(s==k) best=max(best,j-i+1); } }
    return best;
}
int optimal_longest_subarray_sum_k(vector<int>& a,int k){
    unordered_map<int,int> first;
    int s=0, best=0;
    first[0]=-1;
    for(int i=0;i<a.size();++i){
        s+=a[i];
        if(first.find(s)==first.end()) first[s]=i;
        if(first.find(s-k)!=first.end()) best=max(best, i-first[s-k]);
    }
    return best;
}
int main(){return 0;}