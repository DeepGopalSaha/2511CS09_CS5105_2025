#include <bits/stdc++.h>
using namespace std;
long long brute_max_product_subarray(vector<int>& a){
    int n=a.size(); long long best=LLONG_MIN;
    for(int i=0;i<n;++i){ long long p=1; for(int j=i;j<n;++j){ p*=a[j]; best=max(best,p); } }
    return best;
}
long long optimal_max_product_subarray(vector<int>& a){
    long long maxf=1, minf=1, ans=LLONG_MIN;
    for(int x:a){
        if(x==0){ maxf=1; minf=1; ans=max(ans,0LL); continue; }
        long long t1 = maxf * x;
        long long t2 = minf * x;
        maxf = max((long long)x, max(t1, t2));
        minf = min((long long)x, min(t1, t2));
        ans = max(ans, maxf);
    }
    return ans;
}
int main(){return 0;}