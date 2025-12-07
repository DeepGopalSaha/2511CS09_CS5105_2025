#include <bits/stdc++.h>
using namespace std;
vector<vector<string>> findLadders(string beginWord, string endWord, vector<string>& wordList) {
    unordered_set<string> dict(wordList.begin(), wordList.end());
    vector<vector<string>> res;
    if(!dict.count(endWord)) return res;
    unordered_map<string, vector<string>> parents;
    queue<string> q;
    q.push(beginWord);
    unordered_map<string,int> depth;
    depth[beginWord]=0;
    int minDepth = INT_MAX;
    while(!q.empty()){
        string w=q.front(); q.pop();
        int d = depth[w];
        if(d>minDepth) break;
        string cur = w;
        for(int i=0;i<cur.size();i++){
            char orig=cur[i];
            for(char c='a';c<='z';++c){
                cur[i]=c;
                if(!dict.count(cur) && cur!=endWord) continue;
                if(!depth.count(cur)){
                    depth[cur]=d+1;
                    q.push(cur);
                    parents[cur].push_back(w);
                } else if(depth[cur]==d+1){
                    parents[cur].push_back(w);
                }
                if(cur==endWord) minDepth=min(minDepth, d+1);
            }
            cur[i]=orig;
        }
    }
    if(!parents.count(endWord) && beginWord!=endWord) return res;
    vector<string> path;
    function<void(const string&)> dfs = [&](const string& word){
        path.push_back(word);
        if(word==beginWord){
            vector<string> tmp = path;
            reverse(tmp.begin(), tmp.end());
            res.push_back(tmp);
        } else {
            for(auto &p: parents[word]) dfs(p);
        }
        path.pop_back();
    };
    if(endWord==beginWord) return {{beginWord}};
    if(depth.count(endWord)) dfs(endWord);
    return res;
}
int main(){ return 0; }
