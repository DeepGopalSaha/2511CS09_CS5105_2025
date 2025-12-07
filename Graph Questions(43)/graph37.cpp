#include <bits/stdc++.h>
using namespace std;
struct DSU{ vector<int> p, r; DSU(int n):p(n),r(n){ iota(p.begin(),p.end(),0); } int find(int a){ return p[a]==a?a:p[a]=find(p[a]); } bool unite(int a,int b){ a=find(a); b=find(b); if(a==b) return false; if(r[a]<r[b]) swap(a,b); p[b]=a; if(r[a]==r[b]) r[a]++; return true; } };
long long kruskal(int V, vector<array<int,3>>& edges){
    sort(edges.begin(), edges.end(), [](auto &a, auto &b){ return a[2]<b[2]; });
    DSU d(V); long long total=0;
    for(auto &e: edges) if(d.unite(e[0], e[1])) total += e[2];
    return total;
}
int main(){ return 0; }
