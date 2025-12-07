#include <bits/stdc++.h>
using namespace std;
vector<int> brute_rearrange_by_sign(vector<int>& a){
    vector<int> pos, neg;
    for(int x:a){ if(x>0) pos.push_back(x); else neg.push_back(x); }
    vector<int> res;
    for(int i=0;i<pos.size();++i){ res.push_back(pos[i]); res.push_back(neg[i]); }
    return res;
}
vector<int> optimal_rearrange_by_sign(vector<int>& a){
    int n=a.size();
    vector<int> res(n);
    int pi=0, ni=1;
    for(int x:a){
        if(x>0){ res[pi]=x; pi+=2; }
        else { res[ni]=x; ni+=2; }
    }
    return res;
}
int main(){return 0;}