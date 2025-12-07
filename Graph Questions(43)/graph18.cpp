#include <bits/stdc++.h>
using namespace std;
int ladderLength(string beginWord, string endWord, vector<string>& wordList){
    unordered_set<string> dict(wordList.begin(), wordList.end());
    if(!dict.count(endWord)) return 0;
    queue<pair<string,int>>q; q.push({beginWord,1});
    while(!q.empty()){
        auto [word,steps]=q.front(); q.pop();
        if(word==endWord) return steps;
        for(int i=0;i<word.size();i++){
            char orig = word[i];
            for(char c='a'; c<='z'; ++c){
                if(c==orig) continue;
                word[i]=c;
                if(dict.count(word)){ q.push({word, steps+1}); dict.erase(word); }
            }
            word[i]=orig;
        }
    }
    return 0;
}
int main(){ return 0; }
