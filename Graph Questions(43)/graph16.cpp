#include <bits/stdc++.h>
using namespace std;
void solve(vector<vector<char>>& board){
    if(board.empty()) return;
    int n=board.size(), m=board[0].size();
    queue<pair<int,int>>q;
    for(int i=0;i<n;i++) for(int j=0;j<m;j++){
        if((i==0||j==0||i==n-1||j==m-1) && board[i][j]=='O'){
            board[i][j]='T'; q.push({i,j});
        }
    }
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!q.empty()){
        auto [r,c]=q.front(); q.pop();
        for(auto &d:dirs){
            int nr=r+d[0], nc=c+d[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            if(board[nr][nc]=='O'){ board[nr][nc]='T'; q.push({nr,nc}); }
        }
    }
    for(int i=0;i<n;i++) for(int j=0;j<m;j++) if(board[i][j]=='T') board[i][j]='O'; else if(board[i][j]=='O') board[i][j]='X';
}
int main(){ return 0; }
