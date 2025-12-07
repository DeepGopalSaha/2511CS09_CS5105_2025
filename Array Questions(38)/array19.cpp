#include <bits/stdc++.h>
using namespace std;
int brute_longest_subarray_sum_k_positive(vector<int>& a,int k){
    int n=a.size(), best=0;
    for(int i=0;i<n;++i){ int s=0; for(int j=i;j<n;++j){ s+=a[j]; if(s==k) best=max(best,j-i+1); } }
    return best;
}
int optimal_longest_subarray_sum_k_positive(vector<int>& a,int k){
    int n=a.size(), best=0, l=0, s=0;
    for(int r=0;r<n;++r){ s+=a[r]; while(l<=r && s>k){ s-=a[l++]; } if(s==k) best=max(best, r-l+1); }
    return best;
} 
int main(){return 0;}
