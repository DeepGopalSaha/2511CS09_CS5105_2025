#include <bits/stdc++.h>
using namespace std;
struct DSU{ vector<int> p, r; int comps;
    DSU(int n):p(n),r(n),comps(n){ iota(p.begin(),p.end(),0); }
    int find(int a){ return p[a]==a?a:p[a]=find(p[a]); }
    bool unite(int a,int b){ a=find(a); b=find(b); if(a==b) return false; if(r[a]<r[b]) swap(a,b); p[b]=a; if(r[a]==r[b]) r[a]++; comps--; return true; }
};
int makeConnected(int n, vector<pair<int,int>>& edges){
    if((int)edges.size() < n-1) return -1;
    DSU d(n); int extra=0;
    for(auto &e: edges){
        if(!d.unite(e.first, e.second)) extra++;
    }
    int need = d.comps - 1;
    if(extra >= need) return need; else return -1;
}
int main(){ return 0; }
