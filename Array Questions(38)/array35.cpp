#include <bits/stdc++.h>
using namespace std;
int brute_count_subarrays_xor_k(vector<int>& a,int k){
    int n=a.size(), cnt=0;
    for(int i=0;i<n;i++) for(int j=i;j<n;j++){ int x=0; for(int t=i;t<=j;t++) x ^= a[t]; if(x==k) cnt++; }
    return cnt;
}
int optimal_count_subarrays_xor_k(vector<int>& a,int k){
    unordered_map<int,int> pref; pref[0]=1; int xr=0, cnt=0;
    for(int x:a){ xr ^= x; if(pref.count(xr^k)) cnt += pref[xr^k]; pref[xr]++; }
    return cnt;
}
int main(){return 0;}