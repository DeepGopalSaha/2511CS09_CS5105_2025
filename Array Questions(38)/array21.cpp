#include <bits/stdc++.h>
using namespace std;
pair<int,int> brute_two_sum(vector<int>& a,int target){
    int n=a.size();
    for(int i=0;i<n;i++) for(int j=i+1;j<n;j++) if(a[i]+a[j]==target) return {i,j};
    return {-1,-1};
}
pair<int,int> optimal_two_sum(vector<int>& a,int target){
    unordered_map<int,int> mp;
    for(int i=0;i<a.size();++i){
        int need=target-a[i];
        if(mp.count(need)) return {mp[need], i};
        mp[a[i]] = i;
    }
    return {-1,-1};
}
int main(){return 0;}