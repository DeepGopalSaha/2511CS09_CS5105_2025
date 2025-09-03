
#include <bits/stdc++.h>
using namespace std;

void dfstrav(int node, vector<vector<int>> &adj, vector<int> &v, vector<int> &vis, int n){
    vis[node] = 1;
    v.push_back(node);
    
    for(auto it: adj[node]){
        if(!vis[it]){
            dfstrav(it, adj, v, vis, n);
        }
    }
}

vector<int> dfs(vector<vector<int>>& adj) {
    int n = adj.size();
    vector<int> v;
    vector<int> vis(n, 0);
    for(int i = 0; i < n; i++){
        if(!vis[i]) dfstrav(i, adj, v, vis, n); 
    }
    return v;
}

int main() {
    // Example graph: 5 nodes, adjacency list
    vector<vector<int>> adj = {
        {1, 2},   // Node 0 -> 1,2
        {0, 3},   // Node 1 -> 0,3
        {0, 4},   // Node 2 -> 0,4
        {1},      // Node 3 -> 1
        {2}       // Node 4 -> 2
    };

    vector<int> result = dfs(adj);

    cout << "DFS Traversal: ";
    for(int node : result) cout << node << " ";
    cout << endl;
    return 0;
}
