#include <bits/stdc++.h>
using namespace std;
string alienOrder(vector<string>& words, int K){
    vector<set<int>> adj(K);
    vector<int> indeg(K,0);
    for(int i=1;i<words.size();++i){
        string a=words[i-1], b=words[i];
        int len=min(a.size(), b.size());
        int j=0; for(; j<len; ++j){
            if(a[j]!=b[j]){
                if(!adj[a[j]-'a'].count(b[j]-'a')){ adj[a[j]-'a'].insert(b[j]-'a'); indeg[b[j]-'a']++; }
                break;
            }
        }
        if(j==len && a.size()>b.size()) return "";
    }
    queue<int>q; for(int i=0;i<K;i++) if(indeg[i]==0) q.push(i);
    string order;
    while(!q.empty()){
        int u=q.front(); q.pop(); order.push_back('a'+u);
        for(int v: adj[u]) if(--indeg[v]==0) q.push(v);
    }
    if(order.size()!=K) return "";
    return order;
}
int main(){ return 0; }
