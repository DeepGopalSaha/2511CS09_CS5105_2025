#include <bits/stdc++.h>
using namespace std;
int brute_kadane(vector<int>& a){
    int best=INT_MIN;
    int n=a.size();
    for(int i=0;i<n;++i) for(int j=i;j<n;++j){
        int s=0; for(int k=i;k<=j;++k) s+=a[k]; best=max(best,s);
    }
    return best;
}
int optimal_kadane(vector<int>& a){
    int max_ending=INT_MIN, max_so_far=INT_MIN;
    for(int x:a){ if(max_ending==INT_MIN) max_ending=x; else max_ending=max(x, max_ending + x); max_so_far=max(max_so_far, max_ending); }
    return max_so_far;
}
int main(){return 0;}