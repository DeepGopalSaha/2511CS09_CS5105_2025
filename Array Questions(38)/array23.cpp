#include <bits/stdc++.h>
using namespace std;
int brute_kadane(vector<int>& a){
    int n=a.size(), best=INT_MIN;
    for(int i=0;i<n;++i){ int s=0; for(int j=i;j<n;++j){ s+=a[j]; best=max(best,s); } }
    return best;
}
int optimal_kadane(vector<int>& a){
    int max_ending=0, max_so_far=INT_MIN;
    for(int x:a){ max_ending = (max_ending>0? max_ending+x : x); max_so_far = max(max_so_far, max_ending); }
    return max_so_far;
}
int main(){return 0;}