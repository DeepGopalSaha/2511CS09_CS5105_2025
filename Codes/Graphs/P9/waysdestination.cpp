#include <bits/stdc++.h>
using namespace std;

class Solution {
public:
    int countPaths(int n, vector<vector<int>>& roads) {
        // Min-heap priority queue: (distance, node)
        priority_queue<pair<long long,int>, vector<pair<long long,int>>, greater<pair<long long,int>>> pq; 
        int mod=1e9 + 7;
        vector<vector<pair<int,int>>> adj(n);

        // Build adjacency list
        for(auto it: roads){
            adj[it[0]].push_back({it[1],it[2]});
            adj[it[1]].push_back({it[0],it[2]});
        }

        vector<int> ways(n,0);
        vector<long long> dist(n, LLONG_MAX);

        ways[0]=1; 
        dist[0]=0;
        pq.push({0,0});

        while(!pq.empty()){
            auto it = pq.top();
            int node= it.second;
            long long d= it.first;
            pq.pop();

            if (d > dist[node]) continue; 

            for(auto edge: adj[node]){
                int adjnode= edge.first;
                int w= edge.second;
                long long newdist= d + w;

                if(newdist < dist[adjnode]){
                    dist[adjnode] = newdist;
                    pq.push({dist[adjnode], adjnode});
                    ways[adjnode] = ways[node]; // first time reaching
                }
                else if(newdist == dist[adjnode]){
                    ways[adjnode] = (ways[adjnode] + ways[node]) % mod;
                }
            }
        }
        return (ways[n-1]) % mod;
    }
};

// ---------------- MAIN ----------------
int main() {
    Solution sol;

    int n = 7;
    vector<vector<int>> roads = {
        {0,6,7},{0,1,2},{1,2,3},{1,3,3},{6,3,3},
        {3,5,1},{6,5,1},{2,5,1},{0,4,5},{4,6,2}
    };

    cout << sol.countPaths(n, roads) << endl; 
    // Expected output: 4
    return 0;
}
