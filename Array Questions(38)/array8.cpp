#include <bits/stdc++.h>
using namespace std;
int brute_missing_number(vector<int> a){
    int n=a.size();
    for(int x=0;x<=n;++x){
        bool found=false;
        for(int y:a) if(y==x){found=true; break;}
        if(!found) return x;
    }
    return -1;
}
int optimal_missing_number(vector<int>& a){
    int n=a.size();
    long long total = (long long)(n)*(n+1)/2;
    long long sum=0; for(int x:a) sum+=x;
    return (int)(total - sum);
}
int main(){return 0;}