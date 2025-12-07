#include <bits/stdc++.h>
using namespace std;
int brute_best_time_buy_sell(vector<int>& a){
    int n=a.size(), best=0;
    for(int i=0;i<n;++i) for(int j=i+1;j<n;++j) best=max(best, a[j]-a[i]);
    return best;
}
int optimal_best_time_buy_sell(vector<int>& a){
    int minprice=INT_MAX, best=0;
    for(int p:a){ minprice=min(minprice,p); best=max(best, p-minprice); }
    return best;
}
int main(){return 0;}