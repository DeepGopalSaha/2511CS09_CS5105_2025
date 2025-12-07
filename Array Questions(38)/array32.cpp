#include <bits/stdc++.h>
using namespace std;
vector<pair<int,int>> brute_merge_intervals(vector<pair<int,int>>& intervals){
    if(intervals.empty()) return {};
    sort(intervals.begin(), intervals.end());
    vector<pair<int,int>> res; res.push_back(intervals[0]);
    for(int i=1;i<intervals.size();++i){
        auto &last=res.back();
        if(intervals[i].first<=last.second) last.second=max(last.second, intervals[i].second);
        else res.push_back(intervals[i]);
    }
    return res;
}
int main(){return 0;}