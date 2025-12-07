
#include <bits/stdc++.h>
using namespace std;

int rec(int n) {
    if(n==0) return 1;
    if(n<0) return 0;
    return rec(n-1) + rec(n-2);
}

int memo(int n, vector<int>& dp) {
    if(n==0) return 1;
    if(n<0) return 0;
    if(dp[n]!=-1) return dp[n];
    return dp[n] = memo(n-1, dp) + memo(n-2, dp);
}

int tab(int n) {
    vector<int> dp(n+1,0);
    dp[0]=1;
    for(int i=1;i<=n;i++) {
        dp[i] += dp[i-1];
        if(i>=2) dp[i] += dp[i-2];
    }
    return dp[n];
}

int main() {
    return 0;
}
