#include <bits/stdc++.h>
using namespace std;
int brute_count_subarrays_sum_k(vector<int>& a,int k){
    int n=a.size(), cnt=0;
    for(int i=0;i<n;++i) for(int j=i;j<n;++j){ int s=0; for(int t=i;t<=j;++t) s+=a[t]; if(s==k) cnt++; }
    return cnt;
}
int optimal_count_subarrays_sum_k(vector<int>& a,int k){
    unordered_map<int,int> pref; pref[0]=1; int s=0, cnt=0;
    for(int x:a){ s+=x; if(pref.count(s-k)) cnt+=pref[s-k]; pref[s]++; }
    return cnt;
}
int main(){return 0;}