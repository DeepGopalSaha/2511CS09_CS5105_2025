#include <bits/stdc++.h>
using namespace std;
struct DSU{ unordered_map<int,int> p; int find(int x){ if(!p.count(x)) p[x]=x; if(p[x]==x) return x; return p[x]=find(p[x]); } void unify(int a,int b){ a=find(a); b=find(b); if(a!=b) p[b]=a; } };
int removeStones(vector<vector<int>>& stones){
    DSU d; int n=stones.size();
    for(auto &s: stones){
        int r = s[0]; int c = ~s[1];
        d.unify(r, c);
    }
    unordered_map<int,int> compSize;
    for(auto &s: stones) compSize[d.find(s[0])]++;
    int components = compSize.size();
    return n - components;
}
int main(){ return 0; }
