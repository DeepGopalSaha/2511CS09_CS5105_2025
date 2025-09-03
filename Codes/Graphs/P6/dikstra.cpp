#include <iostream>
#include <queue>
#include <vector>
#include <climits>
#include <algorithm>
using namespace std;

class Solution {
  public:
    vector<int> shortestPath(int n, int m, vector<vector<int>>& edges) {
        // adjacency list: adj[node] -> {weight, neighbor}
        vector<pair<int,int>> adj[n+1];
        for (auto it : edges) {
            adj[it[0]].push_back({it[2], it[1]});
            adj[it[1]].push_back({it[2], it[0]});
        }

        vector<int> dist(n+1, INT_MAX);
        vector<int> parent(n+1);

        for (int i = 1; i <= n; i++) parent[i] = i;

        priority_queue<pair<int,int>, vector<pair<int,int>>, greater<pair<int,int>>> pq;

        dist[1] = 0;
        pq.push({0, 1});

        while (!pq.empty()) {
            auto it = pq.top();
            int d = it.first;
            int node = it.second;
            pq.pop();

            for (auto edge : adj[node]) {
                int weight = edge.first;
                int adjnode = edge.second;

                if (d + weight < dist[adjnode]) {
                    dist[adjnode] = d + weight;
                    parent[adjnode] = node;
                    pq.push({dist[adjnode], adjnode});
                }
            }
        }

        if (dist[n] == INT_MAX) return {-1};

        // reconstruct path
        vector<int> path;
        int node = n;
        while (parent[node] != node) {
            path.push_back(node);
            node = parent[node];
        }
        path.push_back(1); // source
        reverse(path.begin(), path.end());

        // answer: first dist[n], then path
        vector<int> ans;
        ans.push_back(dist[n]);
        ans.insert(ans.end(), path.begin(), path.end());
        return ans;
    }
};

int main() {
    int n, m;
    cin >> n >> m;

    vector<vector<int>> edges(m, vector<int>(3));
    for (int i = 0; i < m; i++) {
        cin >> edges[i][0] >> edges[i][1] >> edges[i][2];
    }

    Solution sol;
    vector<int> result = sol.shortestPath(n, m, edges);

    for (int x : result) cout << x << " ";
    cout << endl;

    return 0;
}
