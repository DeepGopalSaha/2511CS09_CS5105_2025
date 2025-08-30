#include<iostream>
#include<vector>
#include<queue>

using namespace std;

vector<int> bfs(vector<vector<int>> adj, int v){
  vector<int> vis(v,0);
  vector<int> ans;
  queue<int> q;

  vis[0]=1;
  q.push(0);

  while(!q.empty()){
    int curr= q.front();
    q.pop();
    ans.push_back(curr);

    for(auto it: adj[curr]){
      if(vis[it]!=1){
        q.push(it);
        vis[it]=1;
      }
    }
  }
  return ans;
}


int main() {
    int v = 5; // number of vertices
    vector<vector<int>> adj(v);

    // undirected graph edges
    adj[0] = {1, 2};
    adj[1] = {0, 3};
    adj[2] = {0, 4};
    adj[3] = {1};
    adj[4] = {2};

    vector<int> traversal = bfs(adj, v);

    cout << "BFS Traversal: ";
    for (int node : traversal) {
        cout << node << " ";
    }
    cout << endl;

    return 0;
}
