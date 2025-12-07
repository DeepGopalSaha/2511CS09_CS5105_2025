#include <bits/stdc++.h>
using namespace std;
struct DSU{ vector<int> p, sz; int count=0; DSU(int n):p(n,-1),sz(n,1){} int find(int x){ return p[x]==x?x:(p[x]=find(p[x])); } void setOn(int x){ if(p[x]==-1){ p[x]=x; count++; } } bool unite(int a,int b){ if(p[a]==-1||p[b]==-1) return false; a=find(a); b=find(b); if(a==b) return false; if(sz[a]<sz[b]) swap(a,b); p[b]=a; sz[a]+=sz[b]; count--; return true; } };
vector<int> numIslands2(int n, int m, vector<pair<int,int>>& ops){
    DSU d(n*m);
    vector<int> ans; int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    for(auto &op: ops){
        int r=op.first, c=op.second, id=r*m+c;
        if(d.p[id]!=-1){ ans.push_back(d.count); continue; }
        d.setOn(id);
        for(auto &dir: dirs){
            int nr=r+dir[0], nc=c+dir[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            int nid=nr*m+nc;
            if(d.p[nid]!=-1) d.unite(id, nid);
        }
        ans.push_back(d.count);
    }
    return ans;
}
int main(){ return 0; }
