#include <bits/stdc++.h>
using namespace std;
vector<vector<int>> brute_three_sum(vector<int>& a){
    int n=a.size(); set<vector<int>> s;
    for(int i=0;i<n;i++) for(int j=i+1;j<n;j++) for(int k=j+1;k<n;k++) if(a[i]+a[j]+a[k]==0){
        vector<int> t={a[i],a[j],a[k]}; sort(t.begin(), t.end()); s.insert(t);
    }
    vector<vector<int>> res(s.begin(), s.end()); return res;
}
vector<vector<int>> optimal_three_sum(vector<int>& a){
    sort(a.begin(), a.end()); int n=a.size();
    vector<vector<int>> res;
    for(int i=0;i<n;i++){
        if(i>0 && a[i]==a[i-1]) continue;
        int l=i+1, r=n-1;
        while(l<r){
            int s=a[i]+a[l]+a[r];
            if(s==0){ res.push_back({a[i],a[l],a[r]}); int val=a[l]; while(l<n && a[l]==val) l++; val=a[r]; while(r>0 && a[r]==val) r--; }
            else if(s<0) l++; else r--;
        }
    }
    return res;
}
int main(){return 0;}