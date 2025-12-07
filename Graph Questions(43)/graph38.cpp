#include <bits/stdc++.h>
using namespace std;
struct DSU{ vector<int> p; DSU(int n):p(n){ iota(p.begin(),p.end(),0); } int find(int a){ return p[a]==a?a:p[a]=find(p[a]); } void unite(int a,int b){ a=find(a); b=find(b); if(a!=b) p[b]=a; } };
vector<vector<string>> accountsMerge(vector<vector<string>>& accounts){
    int n=accounts.size();
    DSU d(n);
    unordered_map<string,int> emailToIdx;
    for(int i=0;i<n;i++){
        for(int j=1;j<accounts[i].size();j++){
            string &e = accounts[i][j];
            if(emailToIdx.count(e)) d.unite(i, emailToIdx[e]);
            else emailToIdx[e]=i;
        }
    }
    unordered_map<int, set<string>> merged;
    for(auto &p: emailToIdx) merged[d.find(p.second)].insert(p.first);
    vector<vector<string>> res;
    for(auto &kv: merged){
        vector<string> cur; cur.push_back(accounts[kv.first][0]);
        for(auto &email: kv.second) cur.push_back(email);
        res.push_back(cur);
    }
    return res;
}
int main(){ return 0; }
