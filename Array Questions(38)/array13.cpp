#include <bits/stdc++.h>
using namespace std;
vector<int> brute_leaders(vector<int>& a){
    int n=a.size(); vector<int> res;
    for(int i=0;i<n;++i){ bool ok=true; for(int j=i+1;j<n;++j) if(a[j]>=a[i]){ok=false; break;} if(ok) res.push_back(a[i]); }
    return res;
}
vector<int> optimal_leaders(vector<int>& a){
    int n=a.size(); vector<int> res; int mx=INT_MIN;
    for(int i=n-1;i>=0;--i){ if(a[i]>mx){ res.push_back(a[i]); mx=a[i]; } }
    reverse(res.begin(), res.end()); return res;
}
int main(){return 0;}