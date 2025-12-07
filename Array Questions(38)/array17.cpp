#include <bits/stdc++.h>
using namespace std;
vector<int> brute_union_sorted(vector<int>& a, vector<int>& b){
    set<int> s(a.begin(), a.end()); for(int x:b) s.insert(x);
    vector<int> res(s.begin(), s.end()); return res;
}
vector<int> optimal_union_sorted(vector<int>& a, vector<int>& b){
    vector<int> res; int i=0,j=0,n=a.size(),m=b.size();
    while(i<n && j<m){
        if(a[i]==b[j]){ if(res.empty()||res.back()!=a[i]) res.push_back(a[i]); i++; j++; }
        else if(a[i]<b[j]){ if(res.empty()||res.back()!=a[i]) res.push_back(a[i]); i++; }
        else { if(res.empty()||res.back()!=b[j]) res.push_back(b[j]); j++; }
    }
    while(i<n){ if(res.empty()||res.back()!=a[i]) res.push_back(a[i]); i++; }
    while(j<m){ if(res.empty()||res.back()!=b[j]) res.push_back(b[j]); j++; }
    return res;
}
int main(){return 0;}