#include <bits/stdc++.h>
using namespace std;
int brute_longest_consecutive(vector<int>& a){
    if(a.empty()) return 0;
    sort(a.begin(), a.end());
    a.erase(unique(a.begin(), a.end()), a.end());
    int best=1, cur=1;
    for(int i=1;i<a.size();++i){
        if(a[i]==a[i-1]+1) cur++; else cur=1;
        best=max(best,cur);
    }
    return best;
}
int optimal_longest_consecutive(vector<int>& a){
    unordered_set<int> s(a.begin(), a.end());
    int best=0;
    for(int x:s){
        if(!s.count(x-1)){
            int cur=x+1;
            while(s.count(cur)) cur++;
            best=max(best, cur-x);
        }
    }
    return best;
}
int main(){return 0;}