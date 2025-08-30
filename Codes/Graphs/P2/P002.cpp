#include <iostream>
#include <vector>
using namespace std;

void dfs(int node, vector<vector<int>>& adj, vector<int>& vis, int v) {
    vis[node] = 1;
    for (int j = 0; j < v; j++) {
        if (adj[node][j] == 1 && !vis[j]) {
            dfs(j, adj, vis, v);
        }
    }
}

int numProvinces(vector<vector<int>> adj, int v) {
    vector<int> vis(v, 0);
    int count = 0;

    for (int i = 0; i < v; i++) {
        if (!vis[i]) {
            count++;
            dfs(i, adj, vis, v);
        }
    }
    return count;
}

int main() {
    vector<vector<int>> adj = {
        {1,1,0},
        {1,1,0},
        {0,0,1}
    };

    int v = adj.size();
    cout << "Number of Provinces: " << numProvinces(adj, v) << endl;

    return 0;
}
