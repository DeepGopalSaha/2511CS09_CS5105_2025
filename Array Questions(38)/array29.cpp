#include <bits/stdc++.h>
using namespace std;
vector<int> brute_majority_element_ii(vector<int>& a){
    unordered_map<int,int> cnt; int n=a.size();
    for(int x:a) cnt[x]++;
    vector<int> res;
    for(auto &p:cnt) if(p.second>n/3) res.push_back(p.first);
    return res;
}
vector<int> optimal_majority_element_ii(vector<int>& a){
    int n=a.size();
    int c1=0,c2=0,v1=0,v2=1;
    for(int x:a){
        if(x==v1) c1++;
        else if(x==v2) c2++;
        else if(c1==0){ v1=x; c1=1; }
        else if(c2==0){ v2=x; c2=1; }
        else{ c1--; c2--; }
    }
    vector<int> res; int cnt1=0,cnt2=0;
    for(int x:a){ if(x==v1) cnt1++; if(x==v2) cnt2++; }
    if(cnt1>n/3) res.push_back(v1);
    if(v2!=v1 && cnt2>n/3) res.push_back(v2);
    return res;
}
int main(){return 0;}