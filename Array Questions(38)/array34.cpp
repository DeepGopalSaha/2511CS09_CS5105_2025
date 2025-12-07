#include <bits/stdc++.h>
using namespace std;
vector<vector<int>> brute_four_sum(vector<int>& a,int target){
    int n=a.size(); set<vector<int>> s;
    for(int i=0;i<n;i++) for(int j=i+1;j<n;j++) for(int k=j+1;k<n;k++) for(int l=k+1;l<n;l++){
        if(a[i]+a[j]+a[k]+a[l]==target){ vector<int> t={a[i],a[j],a[k],a[l]}; sort(t.begin(), t.end()); s.insert(t); }
    }
    return vector<vector<int>>(s.begin(), s.end());
}
vector<vector<int>> optimal_four_sum(vector<int>& a,int target){
    sort(a.begin(), a.end()); int n=a.size();
    vector<vector<int>> res;
    for(int i=0;i<n;i++){
        if(i>0 && a[i]==a[i-1]) continue;
        for(int j=i+1;j<n;j++){
            if(j>i+1 && a[j]==a[j-1]) continue;
            int l=j+1, r=n-1;
            while(l<r){
                long long sum = (long long)a[i]+a[j]+a[l]+a[r];
                if(sum==target){ res.push_back({a[i],a[j],a[l],a[r]}); int lv=a[l], rv=a[r]; while(l<r && a[l]==lv) l++; while(l<r && a[r]==rv) r--; }
                else if(sum<target) l++; else r--;
            }
        }
    }
    return res;
}
int main(){return 0;}