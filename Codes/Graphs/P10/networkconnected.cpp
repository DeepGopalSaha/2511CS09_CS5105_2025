
#include <bits/stdc++.h>
using namespace std;

class DisjointSet{
public:
    vector<int> parent, size, rank;

    DisjointSet(int n){
        parent.resize(n+1);
        size.resize(n+1);
        rank.resize(n+1,0);
        for(int i=0;i<=n;i++){
            parent[i]=i;
            size[i]=1;
        }
    }

    int findUltParent(int node){
        if(parent[node]==node) return node;
        return parent[node] = findUltParent(parent[node]); // path compression
    }

    void unionByRank(int u, int v){
        int ulp_u= findUltParent(u);
        int ulp_v= findUltParent(v);

        if(ulp_u == ulp_v) return;

        if(rank[ulp_u] < rank[ulp_v]){
            parent[ulp_u]= ulp_v;
        }
        else if(rank[ulp_u] > rank[ulp_v]){
            parent[ulp_v]= ulp_u;
        }else{
            parent[ulp_v]= ulp_u;
            rank[ulp_u]++;
        }
    }

    void unionBySize(int u, int v){
        int ulp_u= findUltParent(u);
        int ulp_v= findUltParent(v);

        if(ulp_u == ulp_v) return;

        if(size[ulp_u] < size[ulp_v]){
            parent[ulp_u]= ulp_v;
            size[ulp_v] += size[ulp_u];
        }
        else{
            parent[ulp_v] = ulp_u;
            size[ulp_u] += size[ulp_v];
        }
    }
};

class Solution {
public:
    int makeConnected(int n, vector<vector<int>>& connections) {
        DisjointSet ds(n);
        int extra=0;

        for(auto it: connections){
            int u= it[0];
            int v= it[1];

            if(ds.findUltParent(u) == ds.findUltParent(v)) {
                extra++; // extra cable
            } else {
                ds.unionBySize(u,v);
            }
        }

        int comp=0;
        for(int i=0;i<n;i++){
            if(ds.parent[i]==i) comp++;
        }

        if(extra >= comp-1) return comp-1;
        return -1;
    }
};

// ---------------- MAIN ----------------
int main() {
    Solution sol;

    int n = 6;
    vector<vector<int>> connections = {
        {0,1}, {0,2}, {0,3}, {1,2}, {1,3}
    };

    cout << sol.makeConnected(n, connections) << endl;
    // Expected output: 2 (need 2 operations to connect all 6 computers)

    return 0;
}
