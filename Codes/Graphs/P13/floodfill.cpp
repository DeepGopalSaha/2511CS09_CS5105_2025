
#include <bits/stdc++.h>
using namespace std;

void dfs(vector<vector<int>>& ans, int r, int c, int inicol, int color, int n,int m){
    ans[r][c]= color;
    int delrow[]={-1,0,1,0};
    int delcol[]={0,1,0,-1};

    for(int i=0;i<4;i++){
        int newr=r+delrow[i];
        int newc=c+delcol[i];

        if(newr>=0 && newr<n && newc>=0 && newc<m && ans[newr][newc]== inicol){
            dfs(ans,newr,newc, inicol, color,n,m);
        }
    }
}

void bfs(vector<vector<int>>& ans, int sr, int sc, int inicol, int color, int n, int m){
    queue<pair<int,int>> q;
    q.push({sr,sc});

    int delrow[]={-1,0,1,0};
    int delcol[]={0,1,0,-1};
    while(!q.empty()){
        auto [r,c] = q.front();
        q.pop();

        for(int i=0;i<4;i++){
            int newr=r+delrow[i];
            int newc=c+delcol[i];

            if(newr>=0 && newr<n && newc>=0 && newc<m && ans[newr][newc]== inicol){
                ans[newr][newc]= color;
                q.push({newr,newc});
            }
        }
    }
}

vector<vector<int>> floodFill(vector<vector<int>>& image, int sr, int sc, int color) {
    int inicol= image[sr][sc], n= image.size(), m=image[0].size();
    vector<vector<int>> ans= image;

    if(inicol == color) return image;
    ans[sr][sc]=color;

    // bfs(ans,sr,sc,inicol,color,n,m); // Uncomment for BFS
    dfs(ans,sr,sc,inicol,color,n,m);   // DFS

    return ans;
}

int main() {
    vector<vector<int>> image = {{1,1,1},{1,1,0},{1,0,1}};
    int sr = 1, sc = 1, color = 2;

    vector<vector<int>> result = floodFill(image, sr, sc, color);

    cout << "Flood Fill Result:\n";
    for(auto row : result){
        for(auto val : row) cout << val << " ";
        cout << endl;
    }
    return 0;
}
